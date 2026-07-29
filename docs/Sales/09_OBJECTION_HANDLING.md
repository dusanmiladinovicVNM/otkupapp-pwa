# AgriX Objection Handling Playbook

**Status:** DRAFT v1 — VALIDATION  
**Datum:** 29.07.2026.  
**Svrha:** Standardizovana dijagnostika, obrada i evidencija prigovora bez defanzivnog prodavanja, manipulacije, improvizovanih obećanja ili preventivnog snižavanja cene.

---

## 1. Osnovni princip

Prigovor nije rečenica koju treba pobediti. Prigovor je signal da kupac još nema dovoljno sigurnosti, prioriteta, dokaza, interne saglasnosti ili razloga da pređe na sledeću odluku.

Zato se na prigovor ne odgovara odmah argumentom. Prvo se utvrđuje:

1. šta je sagovornik tačno rekao;
2. šta je stvarni uzrok;
3. koliko je prigovor važan;
4. ko ga još deli;
5. koji dokaz ili odluka može da ga razreši;
6. da li prigovor vraća priliku u discovery, tehničku validaciju, scope review, negotiation ili no-deal.

---

## 2. A-C-T-I-O-N okvir

Svaki ozbiljan prigovor obrađuje se kroz isti tok.

### A — Acknowledge

Priznati legitimnost pitanja bez automatskog slaganja.

> Razumem zašto vam je to važno.

### C — Clarify

Razjasniti šta prigovor stvarno znači.

> Kada kažete da je skupo, da li je problem ukupni budžet, odnos cene i vrednosti ili obim koji smo uključili?

### T — Test

Proveriti da li je to glavni razlog zastoja.

> Kada bismo ovo pitanje razjasnili, postoji li još nešto što bi sprečavalo naredni korak?

### I — Isolate cause

Odvojiti simptom od stvarnog uzroka.

> Da li je veća briga sama tehnologija ili rizik da ljudi ne prihvate novi način rada u sezoni?

### O — Offer evidence or path

Ponuditi odgovarajući dokaz, validaciju ili promenu obuhvata — ne generički argument.

### N — Next decision

Dogovoriti konkretnu sledeću odluku.

> Ako tehnički potvrdimo ovaj tok do petka, možemo li zatim zajedno potvrditi konačni scope?

---

## 3. Kategorije prigovora

| Kategorija | Tipična poruka | Stvarni mogući uzrok | Sledeći proces |
|---|---|---|---|
| Vrednost | „Skupo je“ | vrednost nije dovoljno jasna, scope je preširok | discovery / scope review |
| Budžet | „Nemamo budžet“ | nema odobrenih sredstava ili prioritet nije dovoljno visok | buying process / nurture |
| Status quo | „Excel nam radi“ | trošak promene deluje veći od troška problema | consequence discovery |
| Konkurencija | „Već imamo ERP“ | nejasna uloga AgriX-a u postojećoj arhitekturi | fit / integration review |
| Rizik | „Šta ako sistem stane?“ | strah od operativnog prekida | technical validation |
| Implementacija | „Nemamo vremena za uvođenje“ | sezonski timing ili nedostatak internog vlasnika | implementation scope |
| Autoritet | „Moram da pitam partnera“ | sagovornik nije ekonomski kupac | buying committee map |
| Poverenje | „Vi ste mala firma“ | vendor continuity i podrška nisu dokazani | risk plan / proof |
| Funkcionalni gap | „Nama treba X“ | stvarni must-have ili usputna želja | gap classification |
| Odlaganje | „Javite se pred sezonu“ | pravi timing ili ljubazno odbijanje | trigger-based nurture |
| Interni razvoj | „Imamo svog programera“ | kontrola, sunk cost ili stvarna alternativa | build-vs-buy discovery |
| No decision | „Razmislićemo“ | nema odluke, nema prioriteta ili skriveni veto | decision discovery |

---

## 4. Pravila dijagnostike

Pre odgovora mora biti poznato najmanje:

- ko je izneo prigovor;
- u kojoj fazi procesa;
- da li je prigovor nov ili je ranije bio poznat;
- da li je individualan ili zajednički;
- da li je blocker ili samo pitanje;
- koji dokazni prag sagovornik traži;
- šta će se promeniti ako se prigovor razreši.

Zabranjeno je pretpostaviti da ista rečenica kod dva kupca ima isti uzrok.

---

## 5. Cena i vrednost

### 5.1 „Skupo je“

Prvi odgovor:

> Razumem. Da ne bih odgovorio pogrešno — da li je problem ukupni iznos, odnos cene i koristi ili obim koji smo uključili?

Dijagnostička pitanja:

- U odnosu na koju alternativu je skupo?
- Koji deo ponude vam deluje najmanje opravdano?
- Da li biste isto mislili kada bi prioritetni proces bio potvrđen kao fit?
- Da li je problem investicija sada ili godišnji trošak?
- Ko još ocenjuje ekonomsku opravdanost?

Dozvoljeni odgovori:

- ponovna veza sa potvrđenim problemom i posledicama;
- razdvajanje obaveznog i opcionog scope-a;
- zajednička provera ekonomskog modela;
- fazno uvođenje kada je operativno legitimno;
- poređenje sa troškom statusa quo samo uz validirane podatke kupca.

Nedozvoljeno:

- preventivni popust;
- izmišljeni ROI;
- „to će vam se sigurno isplatiti“;
- napad na jeftiniju alternativu;
- snižavanje cene bez promene obuhvata ili objavljene cenovne strukture.

### 5.2 „Pošaljite samo cenu“

> Mogu da pošaljem važeći cenovnik. Da vam ne pošaljem pogrešnu konfiguraciju, treba mi još da potvrdimo broj stanica, pravnih lica i procese koje želite da obuhvatite. Možemo li to proći za deset minuta?

Ako odbija bilo kakvu kvalifikaciju, šalje se javni cenovnik uz CRM status `PRICE SHOPPING / UNQUALIFIED`, ne individualna ponuda.

### 5.3 „Nemamo budžet“

> Da li to znači da sredstva nisu planirana, da je iznos iznad raspoloživog ili da problem trenutno nije dovoljno visoko na listi prioriteta?

Ishodi:

- nije planirano, ali postoji trigger → nurture sa tačnim datumom;
- prioritet nije dovoljno visok → vratiti u impact discovery;
- nema realne sposobnosti kupovine → zatvoriti ili dugoročni nurture;
- scope je preširok → scope review bez urušavanja osnovnog ishoda.

---

## 6. Status quo i postojeći sistemi

### 6.1 „Excel nam radi“

> Moguće je da za vaš obim zaista radi dovoljno dobro. Gde se danas ipak pojavljuje najviše ručnog usklađivanja, čekanja ili zavisnosti od jedne osobe?

Cilj nije dokazati da je Excel loš. Cilj je utvrditi da li status quo proizvodi problem vredan promene.

No-deal signal:

- nema značajnog problema;
- nema posledice;
- nema planiranog rasta ili regulatornog zahteva;
- kupac ne vidi razlog za promenu.

### 6.2 „Već imamo ERP“

> To samo po sebi nije konflikt. Važno je da razumemo gde ERP danas počinje i završava, a gde nastaju podaci sa otkupnog mesta. Koji deo procesa i dalje ostaje van njega?

Mogući ishodi:

- AgriX kao operativni sistem uz knjigovodstveni ERP;
- integraciona ili eksportna validacija;
- postojeći ERP već rešava prioritetni tok → no-fit;
- kupac očekuje potpunu ERP zamenu → novi discovery i scope.

### 6.3 „Imamo svog programera“

> To može biti dobra prednost. Da li je odluka da se sistem razvija interno već doneta ili još upoređujete trošak, vreme i rizik internog razvoja sa gotovim vertikalnim sistemom?

Obavezno proveriti:

- šta već postoji;
- koliko je razvoj zavisan od jedne osobe;
- ko održava sistem u sezoni;
- ko vodi regulatorne i procesne promene;
- koji je realan rok;
- šta će biti kriterijum odluke build-vs-buy.

---

## 7. Implementacija i operativni rizik

### 7.1 „Nemamo vremena za uvođenje“

> Da li je problem kalendarski period, raspoloživost ljudi ili procena da bi promena ugrozila sezonu?

Mogući putevi:

- pomeranje implementacije uz definisan trigger;
- pre-season priprema;
- fazno uvođenje ograničenog scope-a;
- pilot samo ako ima jasan cilj, kriterijume i vlasnika;
- no-deal ako kupac nema internog vlasnika ni realan termin.

### 7.2 „Ljudi to neće koristiti“

> Koja grupa je najkritičnija i šta je kod prethodnih promena izazvalo otpor?

Proveriti:

- komplikovanost trenutnog rada;
- digitalnu pismenost;
- uređaje i konekciju;
- ko menja rutinu;
- ko gubi status, kontrolu ili neformalnu ulogu;
- obuku i podršku;
- fallback proceduru.

Odgovor nije „lako je za korišćenje“, već konkretna validacija sa stvarnim korisnicima.

### 7.3 „Šta ako internet ne radi?“

Ne davati univerzalan odgovor bez potvrde konkretnog workflow-a i trenutnih mogućnosti proizvoda.

> Hajde da preciziramo na kom mestu i u kom koraku je prekid veze kritičan. Zatim ćemo potvrditi tačno ponašanje sistema i operativnu fallback proceduru za taj scenario.

### 7.4 „Šta ako sistem stane u sezoni?“

> To je legitimna briga. Treba da razdvojimo prevenciju, monitoring, podršku, oporavak i ručnu fallback proceduru. Za svaki kritični proces ćemo navesti odgovornost i očekivano ponašanje.

Ne obećavati apsolutnu dostupnost, nulti rizik ili rok podrške koji nije ugovorno potvrđen.

---

## 8. Poverenje u dobavljača

### 8.1 „Vi ste mala firma“

> Tačno je da AgriX nije velika korporacija. Važno je zato da ne tražimo poverenje na osnovu veličine, već da transparentno pokažemo proizvod, način održavanja, podršku, odgovornosti i plan kontinuiteta.

Dokazi koji mogu biti relevantni:

- stabilnost postojećeg codebase-a;
- update i migration mehanizmi;
- dokumentovana implementacija i podrška;
- ugovorne obaveze;
- reference uz dozvolu;
- plan kontinuiteta i vlasništvo nad podacima;
- jasne granice onoga što se obećava.

### 8.2 „Šta ako prestanete da radite?“

Ovo zahteva konkretan vendor-continuity odgovor usklađen sa ugovorom i tehničkom arhitekturom. Ne improvizovati escrow, prenos izvornog koda ili trajnu podršku ako to nije formalno ponuđeno.

> Razumem. Pripremićemo tačan odgovor kroz ugovorne obaveze, pristup podacima, lokalne komponente i plan kontinuiteta, bez usmenih obećanja van ugovora.

### 8.3 „Nemate dovoljno referenci“

> To je fer primedba. Nećemo nadoknađivati manjak referenci tvrdnjama. Možemo pokazati ono što je dokazivo, definisati ograničenu validaciju i unapred dogovoriti kriterijume na osnovu kojih ćete oceniti rizik.

---

## 9. Funkcionalni i tehnički prigovori

### 9.1 „Nama treba funkcija X“

Pitanja:

- Koji proces bez nje ne može da se završi?
- Koliko često se koristi?
- Ko je korisnik?
- Da li je regulatorni, operativni ili preferencijalni zahtev?
- Postoji li prihvatljiv workaround?
- Da li je must-have pre odluke ili posle prve faze?

Klasifikacija:

- `CORE GAP` — blokira poslovni ishod;
- `COMPLIANCE GAP` — blokira usklađenost;
- `INTEGRATION GAP` — blokira tok sa drugim sistemom;
- `WORKFLOW GAP` — proces je moguć, ali neprihvatljiv;
- `PREFERENCE` — želja bez ozbiljnog poslovnog uticaja;
- `FUTURE` — nije potrebno za početni scope;
- `OUT OF SCOPE` — nije deo proizvoda/ponude.

### 9.2 „Možete li to samo brzo da dodate?“

> Mogu da zabeležim zahtev, ali neću obećati rok bez analize. Prvo treba da potvrdimo poslovnu važnost, zavisnosti, uticaj na standardni proizvod, cenu i implementacioni plan.

### 9.3 „Treba nam integracija sa X“

Obavezno utvrditi:

- smer i učestalost razmene;
- podatke i format;
- vlasnika API-ja ili fajla;
- autentikaciju;
- error handling;
- odgovornost za promene;
- test okruženje;
- kriterijume prihvatanja.

„Može integracija“ nije dozvoljen odgovor bez tehničke procene.

---

## 10. Autoritet, politika i odlaganje

### 10.1 „Moram da pitam partnera/direktora“

> Naravno. Koja pitanja će njemu biti najvažnija i da li ima smisla da ih zajedno prođemo kako ne biste morali da prenosite tehničke i komercijalne detalje?

Ako nema pristupa donosiocu odluke, prilika ne napreduje u forecast-u.

### 10.2 „Pošaljite ponudu pa ćemo razmisliti“

> Ponudu mogu da pripremim kada potvrdimo scope i proces odluke. Predlažem da odmah zakažemo kratki zajednički pregled, da dokument ne ostane bez konteksta i da odmah označimo otvorena pitanja.

### 10.3 „Javite se pred sezonu“

> Razumem. Da li tada planirate stvarnu odluku ili samo želite da ponovo procenite situaciju? Koji događaj ili datum bi bio pravi signal za nastavak?

CRM mora sadržati:

- datum;
- trigger;
- razlog odlaganja;
- kontakt;
- očekivanu odluku pri reaktivaciji.

Bez triggera, status nije aktivna prilika.

### 10.4 „Razmislićemo“

> Naravno. Da bih znao da li ima smisla da ostanemo aktivni: o čemu konkretno treba da razmislite, ko učestvuje i kada očekujete odluku?

Ako nema odgovora, prilika ide u `NO DECISION` ili nurture, ne ostaje u aktivnom forecast-u.

---

## 11. Konkurencija

### 11.1 „Drugi je jeftiniji“

> Moguće. Da li upoređujemo isti obim, iste odgovornosti, implementaciju, podršku i rizik? Hajde da napravimo neutralnu matricu kriterijuma pre nego što razgovaramo samo o ukupnoj ceni.

Ne omalovažavati konkurenciju i ne iznositi neproverene tvrdnje.

### 11.2 „Drugi ima funkciju koju vi nemate“

> To može biti presudno ako je ta funkcija vezana za vaš prioritetni proces. Hajde da proverimo njenu poslovnu važnost i da budemo jasni da li AgriX ima fit, gap ili no-fit.

### 11.3 „Poznajemo njihov tim duže“

To je prigovor poverenja, ne funkcionalnosti.

> Razumem. Šta bi AgriX morao konkretno da dokaže da bi vendor rizik bio prihvatljiv, bez obzira na funkcionalni fit?

---

## 12. Prigovor ili odbijanje

Prigovor se obrađuje samo kada postoji realna mogućnost da novi dokaz ili odluka promeni ishod.

Odbijanje se poštuje kada kupac jasno kaže:

- da nije zainteresovan;
- da ne želi dalji kontakt;
- da problem nije prioritet;
- da je odluka već doneta i ne postoji review trigger;
- da AgriX nema fit za obavezni zahtev.

Odgovor:

> Razumem. Zatvoriću ovu priliku i neću vas dalje kontaktirati u okviru ove teme. Hvala na jasnom odgovoru.

---

## 13. Red flags

- isti prigovor se vraća posle više validacija;
- sagovornik stalno dodaje nove uslove;
- nema pristupa ekonomskom kupcu;
- traži individualni popust bez promene scope-a;
- traži nerealne ili usmene garancije;
- očekuje neograničen razvoj u standardnoj ceni;
- odbija da definiše odluku i rok;
- koristi demo i ponude samo za benchmarking;
- insistira na funkciji koja je van strateškog pravca proizvoda;
- kupovina zavisi od skrivene političke saglasnosti koju niko ne želi da adresira.

Ovi signali zahtevaju requalification, ne dodatno ubeđivanje.

---

## 14. Zabranjene formulacije

Ne koristiti:

- „To nije problem.“
- „Verujte mi.“
- „Svi naši klijenti su zadovoljni.“
- „To ćemo sigurno brzo dodati.“
- „Konkurencija to ne može.“
- „Ovo je poslednja cena samo danas.“
- „Ako ne odlučite sada, izgubićete priliku.“
- „Sistem nikada ne pada.“
- „Implementacija je jednostavna.“
- „Ljudi će se brzo navići.“
- „To se samo jednom plati i rešili ste sve.“

---

## 15. CRM Objection Record

Za svaki značajan prigovor beleži se:

- datum i faza;
- sagovornik i uloga;
- doslovna formulacija;
- kategorija;
- pretpostavljeni uzrok;
- potvrđeni uzrok;
- poslovna važnost;
- blocker: da/ne;
- ko još deli prigovor;
- dokaz ili akcija potrebna za razrešenje;
- vlasnik akcije;
- rok;
- rezultat;
- uticaj na stage, scope, forecast i close date;
- da li je formulacija ili dokaz kandidovan za playbook reviziju.

---

## 16. Quality score

Svaki ozbiljan objection razgovor ocenjuje se 0–2 po oblasti:

| Oblast | 0 | 1 | 2 |
|---|---|---|---|
| Acknowledgement | defanzivno | neutralno | legitimnost jasno priznata |
| Clarification | pretpostavljen uzrok | delimično | potvrđen stvarni uzrok |
| Isolation | nije provereno | delimično | potvrđeno da li je glavni blocker |
| Evidence | generički argument | relevantan, ali slab | odgovarajući dokaz/validacija |
| Next decision | bez koraka | nejasan korak | vlasnik i datum potvrđeni |
| CRM record | nema | nepotpun | kompletan |

Maksimum: 12.

- 10–12: kvalitetno obrađen prigovor;
- 7–9: potreban coaching;
- 0–6: prigovor verovatno samo potisnut, ne razrešen.

---

## 17. Validacioni plan

Posle prvih 50 značajnih prigovora analizirati:

- učestalost po kategoriji i personi;
- fazu u kojoj se pojavljuju;
- koliko često je prvi iskaz bio stvarni uzrok;
- koji dokazi najčešće menjaju odluku;
- koji prigovori vraćaju priliku u discovery;
- koji vode u no-deal;
- gde se ponavljaju product gap-ovi;
- gde prodajni proces stvara prigovor prekasnim otkrivanjem rizika;
- da li neki email, call ili demo obrazac izaziva nepotreban otpor.

Playbook se menja na osnovu ponovljenih dokaza, ne na osnovu jedne anegdote.

---

## 18. Checklist

Pre odgovora:

- [ ] Znam tačnu formulaciju prigovora.
- [ ] Razdvojio sam simptom od uzroka.
- [ ] Znam da li je blocker.
- [ ] Znam ko ga još deli.
- [ ] Nisam obećao nedokazanu funkciju, rok ili rezultat.

Pre završetka razgovora:

- [ ] Prigovor je razrešen, otvoren sa planom ili potvrđen kao no-fit.
- [ ] Postoji vlasnik i datum sledeće akcije.
- [ ] Stage i forecast su korigovani.
- [ ] CRM record je popunjen.
- [ ] Novi dokaz ili obrazac je označen za Customer Intelligence Loop.
