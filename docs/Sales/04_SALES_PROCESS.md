# 04 — AgriX Sales Process

**Status:** DRAFT v1 — VALIDATION  
**Vlasnik:** osnivač AgriX-a  
**Datum:** 2026-07-28  
**Povezani dokumenti:** `00_COMMERCIAL_OPERATING_SYSTEM_ROADMAP.md`, `01_MARKET_POSITIONING.md`, `02_PSYCHOLOGICAL_PROFILES.md`, `03_BUYING_PROCESS.md`

---

## 1. Svrha

Ovaj dokument definiše ponovljiv način na koji AgriX vodi prodajnu priliku od identifikovanog naloga do potpisanog posla, kontrolisanog odustajanja ili dugoročnog nurture režima.

Sales process nije lista aktivnosti prodavca. Njegova svrha je da kupcu pomogne da završi sopstveni buying process, uz jasan dokaz napretka, kontrolu rizika i dogovoren sledeći korak.

Osnovno pravilo:

> Faza se određuje prema onome što je kupac dokazano završio, a ne prema poslednjoj aktivnosti koju je AgriX obavio.

Demo, ponuda, telefonski poziv i poslati mejl nisu sami po sebi napredak.

---

## 2. Operativni principi

1. **Buyer evidence over seller activity.** Aktivnost je korisna samo ako proizvodi novu kupčevu odluku, proveru ili obavezu.
2. **No next step, no active opportunity.** Svaka aktivna prilika ima sledeći korak, vlasnika i datum.
3. **No discovery, no tailored demo.** Demo bez razumevanja procesa je prezentacija proizvoda, ne prodajni napredak.
4. **No scope, no proposal.** Ponuda se ne šalje dok obim, paketi, moduli, instance, implementacija i otvorene pretpostavke nisu razjašnjeni.
5. **Proposal is presented, not emailed into silence.** Ponuda se predstavlja na sastanku ili pozivu.
6. **No artificial urgency.** Rok mora dolaziti iz sezone, implementacije, poslovnog događaja ili kupčevog internog procesa.
7. **No individual discounts.** Važi komercijalna odluka da nema pregovaračkih ni individualnih popusta.
8. **Disqualify professionally.** Nije svaki nalog prilika; rano i argumentovano odustajanje čuva kapacitet.
9. **One source of truth.** CRM mora sadržati isti stage, decision date, scope, rizike i sledeći korak koji se koriste u forecast-u.
10. **Sales-to-implementation continuity.** Prodaja ne sme obećati nešto što implementacija nije prihvatila.

---

## 3. Struktura pipeline-a

| Faza | Naziv | Buying stage veza | Suština |
|---|---|---|---|
| S0 | Target Account | B0 | nalog odgovara ICP-u, ali nema potvrđenog kontakta ili potrebe |
| S1 | Connected | B0–B1 | uspostavljen relevantan kontakt i dozvoljen nastavak razgovora |
| S2 | Qualified Problem | B1–B2 | potvrđen problem, posledica, owner i vremenski kontekst |
| S3 | Discovery | B2 | mapirani proces, stakeholderi, success criteria i rizici |
| S4 | Solution Evaluation | B3 | AgriX se proverava prema konkretnim procesima i kriterijumima |
| S5 | Risk Alignment | B4 | usaglašavaju se implementacija, migracija, obuka, podrška i tehnički uslovi |
| S6 | Proposal Review | B5 | konačni scope i komercijalni model su predstavljeni relevantnim ljudima |
| S7 | Decision / Approval | B5–B6 | kupac završava interno odobrenje, ugovor i commitment |
| S8 | Closed Won | B6–B7 | postoji formalna odluka i spreman handoff |
| S9 | Closed Lost | — | izabran konkurent, status quo ili drugi eksplicitan ishod |
| SN | Nurture | B0–B2 | postoji fit ili potencijal, ali nema aktivnog projekta ili roka |

`DECISION`: Aktivni forecast obuhvata samo S2–S7. S0, S1 i SN nisu aktivan prihod u forecast-u.

---

# 4. Stage playbook

## S0 — Target Account

### Svrha

Identifikovati naloge koji imaju razuman potencijal za AgriX, bez pretvaranja baze kontakata u lažni pipeline.

### Entry kriterijumi

- firma je identifikovana;
- postoji dokaz ili razumna procena da posluje u organizovanom otkupu ili povezanoj ciljnoj vertikali;
- poznat je osnovni razlog fit-a: stanice, kooperanti, operativni tok, dokumentacija, logistika, lager ili finansijski tok.

### Obavezne aktivnosti

- desk research;
- identifikacija 1–3 relevantne uloge;
- hipoteza o mogućem triggeru;
- izbor prve poruke prema personi i sezonskom trenutku.

### Exit u S1

- uspostavljena dvosmerna komunikacija sa relevantnom osobom;
- osoba dozvoljava nastavak razgovora ili daje relevantnog sagovornika.

### Nije dokaz napretka

- mejl je poslat ili otvoren;
- kontakt je prihvatio LinkedIn zahtev;
- centrala je dala generičku adresu;
- nepoznata osoba je preuzela materijal.

### Maksimalno trajanje

Nema ograničenja kao account list, ali ne ulazi u aktivan pipeline.

---

## S1 — Connected

### Svrha

Utvrditi da li postoji poslovni razlog za kvalifikacioni razgovor.

### Entry kriterijumi

- ostvaren kontakt sa osobom koja poznaje proces ili može da usmeri ka njoj;
- postoji dozvola za kratku razmenu o trenutnom načinu rada.

### Obavezna pitanja

- Šta se promenilo ili šta bi moglo pokrenuti razmatranje novog sistema?
- Koji deo procesa je trenutno najteže kontrolisati?
- Da li postoji određena sezona, projekat ili rok?
- Ko je odgovoran za taj proces?

### Exit u S2

- potvrđen konkretan problem ili aktivni trigger;
- imenovan problem owner;
- potvrđena posledica statusa quo;
- poznat okvirni rok ili sezonski kontekst;
- dogovoren kvalifikacioni/discovery korak sa datumom.

### Prelazak u SN

- nalog ima fit, ali nema aktivni trigger, rok ili spremnost za razgovor;
- kontakt eksplicitno traži da se tema otvori u određenom budućem periodu.

### Closed Lost / Disqualified

- nema relevantnog procesa ili obima;
- traži se samo funkcija koju AgriX ne treba da prodaje samostalno;
- nema pristupa odgovornoj osobi niti realne putanje do nje.

### SLA

- odgovor ili recap istog radnog dana;
- sledeći dogovoreni korak najkasnije u roku od 10 radnih dana, osim eksplicitnog sezonskog datuma.

---

## S2 — Qualified Problem

### Svrha

Dokazati da postoji realan poslovni problem vredan daljeg ulaganja vremena sa obe strane.

### Entry kriterijumi

- problem i posledica su potvrđeni;
- postoji problem owner;
- postoji razlog zašto se tema razmatra sada;
- dogovoren je discovery.

### Minimalna kvalifikacija — PACT

- **Problem:** šta tačno ne funkcioniše ili ne skaluje;
- **Authority path:** ko utiče, ko odobrava, ko može blokirati;
- **Consequence:** operativna, finansijska, regulatorna ili reputaciona posledica;
- **Timing:** decision date i najkasniji smisleni implementation start.

### Obavezna CRM polja

- trigger;
- problem statement rečima kupca;
- posledica;
- problem owner;
- ekonomski kupac — poznat/nepoznat;
- okvirni broj stanica, korisnika i firmi;
- postojeći sistemi;
- decision date;
- next step, owner, date.

### Exit u S3

- discovery agenda prihvaćena;
- uključene najmanje dve relevantne funkcije ili postoji jasan plan za njihovo uključivanje;
- kupac pristaje da opiše stvarni proces, izuzetke i kriterijume uspeha.

### Stagnation pravilo

Ako discovery nije zakazan u 15 kalendarskih dana i ne postoji opravdan sezonski razlog, prilika prelazi u SN ili se zatvara.

---

## S3 — Discovery

### Svrha

Razumeti trenutni proces dovoljno dobro da se utvrdi fit, poslovna vrednost, rizik promene i način evaluacije.

### Entry kriterijumi

- kvalifikovan problem;
- poznata agenda;
- prisutni relevantni procesni sagovornici.

### Obavezni discovery ishodi

1. current-state mapa od nastanka podatka do centralne kontrole;
2. tri najvažnija problema i njihove posledice;
3. kritični izuzeci;
4. desired future state;
5. 3–5 success criteria;
6. stakeholder mapa;
7. status quo koristi i rizik promene;
8. postojeći ERP i način koegzistencije;
9. okvirni scope;
10. decision process i rok;
11. dogovoren evaluatorni sledeći korak.

### Exit u S4

- dokumentovan discovery recap potvrđen od kupca;
- poznati success criteria;
- demo/evaluacija imaju unapred definisanu svrhu;
- mapiran buying committee ili jasno označene nepoznate uloge;
- postoji preliminarni fit bez kritične nerešive praznine;
- zakazan demo ili drugi validation step sa datumom.

### Ne sme se preći u S4 ako

- demo se traži samo kao opšti obilazak;
- nema potvrđenog problema;
- success criteria nisu poznati;
- kontakt odbija pristup osobama koje nose procesni ili ekonomski rizik.

### SLA

- discovery recap u roku od jednog radnog dana;
- demo plan i agenda u roku od dva radna dana;
- evaluatorni korak idealno u narednih 7–10 radnih dana.

---

## S4 — Solution Evaluation

### Svrha

Pomoći kupcu da proveri da li AgriX rešava njegove kritične tokove i success criteria.

### Entry kriterijumi

- potvrđen discovery;
- demo/evaluacija su personalizovani;
- poznati prisutni stakeholderi;
- poznato je šta se pokazuje, šta se ne pokazuje i zašto.

### Obavezni elementi

- scenario zasnovan na kupčevom toku;
- standardni tok i najmanje jedan kritični izuzetak;
- veza između funkcionalnosti i poslovnog ishoda;
- otvoreno označene praznine, pretpostavke i ograničenja;
- potvrda načina koegzistencije sa postojećim ERP-om;
- live capture pitanja i rizika;
- dogovor kako će kupac interno proceniti rezultat.

### Exit u S5

- kupac potvrđuje funkcionalni fit za kritične tokove;
- otvorene praznine su klasifikovane: standard, konfiguracija, implementacija, razvoj, nije podržano;
- poznato je ko mora potvrditi tehnički i operativni rizik;
- kupac pristaje na implementation/risk session;
- nema nepoznatog kritičnog veta.

### Negativan ishod

Ako ključni success criterion nije podržan i nema prihvatljivog rešenja, prilika se zatvara ili re-scope-uje. Ne obećava se razvoj bez procene.

### SLA

- demo recap istog ili narednog radnog dana;
- otvorena pitanja imaju vlasnika i rok;
- risk session u roku od 7 radnih dana kada je moguće.

---

## S5 — Risk Alignment

### Svrha

Smanjiti kupčev percipirani i stvarni rizik implementacije pre komercijalne odluke.

### Entry kriterijumi

- funkcionalni fit je potvrđen;
- postoji realna namera da se proceni uvođenje;
- poznati su ključni stakeholderi i otvorena pitanja.

### Obavezno usaglašavanje

- konačni preliminarni scope;
- broj firmi, instanci, stanica, korisnika i uređaja;
- moduli i paket;
- migracija početnih podataka;
- hardverski i mrežni preduslovi;
- obuka;
- odgovornosti kupca i AgriX-a;
- rollout redosled;
- support model;
- fallback i incident escalation;
- standard naspram posebnog razvoja;
- readiness i najkasniji implementation start.

### Exit u S6

- postoji pisani solution/implementation outline;
- svi kritični rizici imaju odgovor, vlasnika ili eksplicitno prihvatanje;
- scope je dovoljno stabilan za ponudu;
- ekonomski kupac ili ovlašćeni sponsor zna da se priprema ponuda;
- dogovoren proposal review sastanak.

### Ne sme se poslati ponuda ako

- ključni scope elementi nisu poznati;
- postoji nerešen kritični funkcionalni ili tehnički fit;
- kontakt samo želi cenovnik za internu komparaciju bez konteksta;
- nije poznato ko odlučuje niti kako će se odluka doneti.

### SLA

- implementation outline u roku od dva radna dana od risk session-a;
- ponuda u roku od tri radna dana nakon potvrde scope-a, osim kompleksnog posebnog razvoja.

---

## S6 — Proposal Review

### Svrha

Predstaviti tačan poslovni i komercijalni dogovor, proveriti razumevanje i identifikovati preostale odluke.

### Entry kriterijumi

- scope potvrđen;
- ponuda usklađena sa važećim cenovnikom;
- poznat approval path;
- zakazan sastanak za predstavljanje ponude.

### Struktura proposal review-a

1. ponoviti trigger i success criteria;
2. potvrditi dogovoreni scope;
3. pokazati šta je uključeno i šta nije;
4. objasniti implementation pristup i odgovornosti;
5. predstaviti cenu i jedinice obračuna;
6. proveriti sva otvorena pitanja;
7. pitati kako će kupac doneti konačnu odluku;
8. dogovoriti Mutual Action Plan do odluke i kickoff-a.

### Exit u S7

- kupac potvrđuje da ponuda odgovara dogovorenom scope-u;
- poznate su preostale osobe i odobrenja;
- poznat je decision date;
- pravna, nabavna ili ugovorna pitanja imaju vlasnika;
- postoji konkretan sledeći korak sa datumom;
- nema prikrivenog zahteva za individualni popust.

### Stagnation pravilo

Ponuda bez feedback-a i bez održanog review-a nije S6; vraća se u S5 ili SN. Ako kupac dva puta ne izvrši dogovoreni decision step bez novog razloga, prilika se označava `AT RISK`.

### SLA

- recap u roku od jednog radnog dana;
- prvi decision follow-up prema dogovorenom datumu, ne proizvoljno;
- bez dogovorenog datuma: najkasnije pet radnih dana, uz zahtev za jasnoću procesa.

---

## S7 — Decision / Approval

### Svrha

Upravljati završnim odobrenjima bez lažnog forecast optimizma.

### Entry kriterijumi

- proposal review završen;
- finalni scope potvrđen;
- postoji approval path i decision date;
- preostali koraci su konkretno imenovani.

### Obavezna CRM polja

- ekonomski kupac;
- formalni potpisnik;
- champion status;
- decision criteria;
- decision process;
- decision date;
- pravni/nabavni status;
- glavni konkurent ili status quo;
- otvoreni rizici;
- MAP koraci;
- forecast kategorija.

### Exit u S8

- prihvaćena ponuda ili potpisan ugovor prema važećem procesu;
- potvrđen datum početka;
- imenovani projektni vlasnici;
- potvrđen prvi onboarding korak;
- prodajni handoff paket kompletan.

### Exit u S9

- kupac izabrao konkurenta;
- eksplicitno izabrao status quo;
- projekat otkazan;
- rok je prošao i implementacija više nema smisla;
- uslovi su neprihvatljivi ili fit nije dovoljan.

### Verbal commit pravilo

Usmeno „dogovoreno“ nije Closed Won i ne ulazi u Commit forecast bez konkretnog završnog koraka, vlasnika i datuma.

---

## S8 — Closed Won

### Obavezni dokazi

- formalno prihvatanje;
- potvrđen komercijalni scope;
- implementation start;
- kontakt osobe i odgovornosti;
- handoff dokument;
- evidentirana prodajna obećanja i otvorena pitanja.

### Handoff paket

- problem statement i trigger;
- success criteria;
- kupljeni paket, moduli, instance i stanice;
- procesni scope;
- konfiguracione odluke;
- migracioni zahtevi;
- stakeholder mapa;
- rizici i osetljive teme;
- posebna obećanja — samo potvrđena;
- rokovi i sezonski kontekst;
- agreed first value;
- reference/case-study permission status.

Prodajni owner ostaje uključen do formalnog kickoff-a i potvrde da implementacija razume kontekst.

---

## S9 — Closed Lost

### Obavezni reason codes

- izabran konkurent;
- status quo / no decision;
- nedovoljan funkcionalni fit;
- rizik implementacije;
- nedostatak internog kapaciteta;
- cena/budžet;
- pogrešan timing;
- izgubili champion-a;
- nema pristupa autoritetu;
- posebni zahtevi van strategije proizvoda;
- projekat otkazan;
- drugi razlog — obavezan opis.

### Win/loss beleška

Mora sadržati:

- stvarni razlog, ne samo deklarisani prigovor;
- fazu u kojoj je ishod postao verovatan;
- dokaz koji je nedostajao;
- competitor/status quo prednost;
- da li i kada postoji smislen reactivation datum;
- jednu preporuku za playbook, proizvod ili positioning.

Closed Lost se ne koristi kao kazna za prodavca. Netočni reason codes uništavaju učenje.

---

## SN — Nurture

### Kada se koristi

- fit postoji, aktivni trigger ne postoji;
- rok je daleko i kupac nije spreman za projekat;
- tema mora biti otvorena u sledećem predsezonskom prozoru;
- sponsor postoji, ali nema budžeta ili organizacionog kapaciteta;
- projekat je privremeno pauziran sa jasnim razlogom.

### Obavezno

- nurture reason;
- relevantna persona;
- sledeći poslovni događaj;
- reactivation datum;
- sledeća poruka ili dokaz koji ima smisla;
- zabrana generičkog mesečnog spamovanja.

Nurture bez datuma i razloga je arhiva, ne proces.

---

# 5. Next-step discipline

Svaki aktivni opportunity mora imati:

- **akciju:** konkretan glagol i rezultat;
- **vlasnika:** AgriX ili imenovana osoba kod kupca;
- **datum:** precizan datum;
- **kupčev commitment:** šta kupac radi, ne samo šta AgriX šalje;
- **stage evidence:** koji exit kriterijum taj korak završava.

Loš next step:

> „Javiti se sledeće nedelje.“

Dobar next step:

> „Operativni direktor do 4. avgusta šalje mapu stanica i spisak procesa; 6. avgusta održavamo discovery sa administratorom i finansijama radi potvrde success criteria.“

Ako kupac ne prihvata nikakvu obavezu, prilika verovatno nije u aktivnoj evaluaciji.

---

# 6. SLA i maksimalna starost faze

| Faza | Standardni SLA za sledeću aktivnost | At-risk prag | Obavezna reakcija |
|---|---:|---:|---|
| S1 Connected | 5 radnih dana | 10 radnih dana | kvalifikovati, nurture ili zatvoriti |
| S2 Qualified Problem | 7 radnih dana | 15 kalendarskih dana | zakazati discovery ili SN |
| S3 Discovery | 5 radnih dana do recap-a/next step-a | 14 dana bez kupčevog commitment-a | re-kvalifikacija |
| S4 Solution Evaluation | 7 radnih dana | 21 dan | identifikovati blocker ili SN/S9 |
| S5 Risk Alignment | 7 radnih dana | 21 dan | executive alignment ili zatvaranje |
| S6 Proposal Review | prema MAP-u, najviše 5 radnih dana bez datuma | 14 dana bez feedback-a | zahtev za jasnu odluku procesa |
| S7 Decision | prema potvrđenom decision date-u | propušten decision date | forecast downgrade i re-kvalifikacija |

Sezonski opravdano odlaganje se dokumentuje kao izuzetak, sa novim datumom. Ne briše se starost prilike bez razloga.

---

# 7. Forecast kategorije

## Pipeline

- S2–S5;
- postoji realan problem i proces evaluacije;
- nisu završeni komercijalni i approval koraci.

## Best Case

- S6–S7;
- scope i vrednost potvrđeni;
- poznat decision process i datum;
- postoji kredibilan champion;
- još postoji jedan ili više materijalnih rizika.

## Commit

Dozvoljeno samo ako su svi uslovi ispunjeni:

- S7;
- ekonomski kupac je uključen ili je odobrenje eksplicitno potvrđeno;
- finalni scope i cena prihvaćeni;
- poznat potpisni/approval korak;
- decision date je u forecast periodu;
- nema nerešenog materijalnog veta;
- Mutual Action Plan se izvršava.

## Nurture

- nije aktivni forecast;
- postoji fit, ali nema aktivnog projekta ili roka.

Forecast se ne zasniva na osećaju, tonu razgovora ili frazi „sviđa nam se“.

---

# 8. Pipeline hygiene

Jednom nedeljno se proverava:

1. da li stage ima validan evidence;
2. da li postoji next step sa datumom;
3. da li je decision date realan;
4. da li je champion test i dalje zadovoljen;
5. da li postoji novi stakeholder ili veto;
6. da li je scope promenjen;
7. da li je starost faze prešla prag;
8. da li forecast kategorija ima dokaze;
9. da li priliku treba vratiti u raniju fazu;
10. da li je profesionalnije zatvoriti je.

Zabranjeno je:

- automatski pomerati decision date u budućnost;
- držati ponudu otvorenu bez kupčevog odgovora;
- označiti demo kao kvalifikaciju;
- unositi punu vrednost nepoznatog scope-a;
- skrivati izgubljenu priliku kao nurture;
- proglašavati kontakt champion-om bez dokaza;
- držati priliku u S7 na osnovu usmenog optimizma.

---

# 9. Re-kvalifikacija i povratak faze

Faza se sme vratiti unazad. To nije neuspeh, već korekcija realnosti.

Obavezna re-kvalifikacija kada:

- odlazi champion ili problem owner;
- menja se ekonomski kupac;
- scope se značajno proširi;
- decision date se pomeri više od jednom;
- pojavi se novi tehnički ili politički veto;
- kupac menja success criteria;
- sezonski prozor postane nerealan;
- projekat prelazi iz jedne firme na grupu firmi.

Posle re-kvalifikacije stage se određuje prema najnižem nezavršenom buying zadatku.

---

# 10. No-deal pravila

AgriX treba profesionalno da odustane ili ne uđe u aktivan proces kada:

- kupac traži trajni klijentski fork;
- traži kritične funkcije koje nisu spremne, uz očekivanje da se obećaju bez procene;
- insistira na individualnom popustu suprotno komercijalnoj odluci;
- nema internog vlasnika implementacije;
- ne želi da uključi ključne korisnike, a očekuje punu odgovornost dobavljača;
- implementation timing objektivno ugrožava sezonu;
- traži da se zaobiđu audit, odgovornost ili zakonska pravila;
- odnos vrednosti i support rizika nije održiv;
- AgriX nije odgovarajuće rešenje.

Profesionalna formulacija:

> „Na osnovu obima i uslova koje smo razjasnili, ne mogu odgovorno da tvrdim da je AgriX sada pravo rešenje za Vas. Bolje je da to kažemo pre implementacije nego tokom sezone. Evidentiraću šta bi moralo da se promeni da bismo temu ponovo otvorili.“

---

# 11. Handoff ka implementaciji

Closed Won nije kraj prodajnog procesa dok nije održan interni handoff.

## Obavezni handoff sastanak

Učesnici:

- sales owner;
- implementation owner;
- po potrebi product/technical owner.

Agenda:

1. zašto je kupac kupio;
2. success criteria;
3. ko je champion, a ko potencijalni blocker;
4. šta je standardni scope;
5. šta je eksplicitno van scope-a;
6. rokovi i sezonski rizik;
7. migracija, oprema i obuka;
8. osetljive političke ili procesne teme;
9. sva obećanja data kupcu;
10. first-value milestone.

Implementation owner ima pravo da zaustavi kickoff ako handoff nije kompletan ili ako otkrije neodobreno obećanje.

---

# 12. CRM minimum za svaki aktivni opportunity

- account i firma;
- opportunity owner;
- stage;
- stage entered date;
- trigger;
- problem statement;
- quantified ili kvalitativna posledica;
- current solution;
- scope summary;
- paket/moduli — preliminarno ili finalno;
- broj firmi, instanci, stanica i korisnika;
- economic buyer;
- champion i champion score;
- buying committee;
- decision criteria;
- decision process;
- decision date;
- implementation deadline;
- competition/status quo;
- top three risks;
- next step, owner, date;
- forecast category;
- amount i confidence basis;
- loss/nurture reason kada je relevantno.

Polje bez informacije se označava `UNKNOWN`, ne popunjava pretpostavkom.

---

# 13. Validacija v1

Na prvih 20 kvalifikovanih prilika pratiti:

- koliko prilika ulazi u S2 bez realnog triggera;
- prosečno trajanje po fazi;
- gde se najčešće pojavljuje skriveni veto;
- procenat demoa bez potvrđenih success criteria;
- procenat ponuda bez proposal review-a;
- broj pomeranja decision date-a;
- tačnost Best Case i Commit forecast-a;
- win/loss/no-decision razloge;
- da li maksimalna starost faza odgovara realnom sezonskom ciklusu;
- da li stage gates ubrzavaju ili nepotrebno komplikuju prodaju.

Posle 10 prilika radi se mini-revizija. Posle 20 prilika izdaje se v2 sa stvarnim konverzijama i median stage duration.

---

# 14. Definition of Done

Sales Process v1 je operativno primenljiv kada:

- CRM koristi iste faze i obavezna polja;
- najmanje pet prilika je vođeno kroz stage gates;
- nijedna aktivna prilika nema next step bez vlasnika i datuma;
- proposal review je standard, ne izuzetak;
- forecast kategorije se dodeljuju po dokazima;
- Closed Won handoff paket koristi implementacija;
- win/loss razlozi se analiziraju mesečno;
- stage aging i stagnation pravila se zaista primenjuju.
