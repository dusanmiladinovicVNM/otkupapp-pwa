# AgriX — psihološki profili kupaca i buying committee

**Status:** DRAFT v1 — za validaciju  
**Vlasnik:** osnivač AgriX-a  
**Datum:** 2026-07-28  
**Povezano:** `00_COMMERCIAL_OPERATING_SYSTEM_ROADMAP.md`, `../Master Plan/03_CUSTOMERS_AND_JOBS.md`, `../Master Plan/02_STRATEGY.md`

---

## 1. Svrha

Ovaj dokument opisuje kako različite uloge u ciljnoj firmi doživljavaju:

- poslovni problem;
- rizik promene i rizik ostanka na postojećem načinu rada;
- ličnu i profesionalnu posledicu odluke;
- dokaz potreban da bi poverovale AgriX-u;
- jezik koji otvara razgovor;
- ponašanje koje ukazuje na stvarnu kupovnu nameru;
- ponašanje koje može blokirati prodaju ili implementaciju.

Ovo nisu opšti karakterološki stereotipi. Profil se odnosi isključivo na ponašanje osobe u kontekstu procene, kupovine i uvođenja AgriX-a.

---

## 2. Metodologija

Profil se gradi iz šest slojeva:

1. **Formalna uloga** — šta osoba zvanično radi.
2. **Operativna odgovornost** — za šta će biti pozvana kada sezona ne ide po planu.
3. **Lični rizik odluke** — šta osoba može izgubiti ako projekat ne uspe.
4. **Status quo nagrada** — šta osoba dobija time što ništa ne menja.
5. **Desired progress** — kakvu promenu želi, čak i kada je ne opisuje kao potrebu za softverom.
6. **Evidence threshold** — koji dokaz mora videti da bi podržala sledeći korak.

Svaka tvrdnja je označena kao `FACT`, `EVIDENCE` ili `HYPOTHESIS`.

---

## 3. Buying committee mapa

| Uloga | Tipična funkcija u odluci | Može pokrenuti | Može blokirati | Najvažniji dokaz |
|---|---|---:|---:|---|
| Vlasnik / generalni direktor | ekonomski kupac i konačni autoritet | da | da | poslovna kontrola, referenca, rizik implementacije, ROI |
| Operativni direktor / rukovodilac otkupa | problem owner i često champion | da | da | realan tok sezone, brzina, pregled mreže, izuzeci |
| Centralni administrator / ključni Excel operator | stručni evaluator i implementacioni čuvar | retko | veoma često | tačnost, kontrola korekcija, dokumentni lanac, smanjenje rada |
| Finansije / knjigovodstvo | evaluator finansijskog i regulatornog toka | ponekad | da | SEF, banka, salda, audit, usklađenost sa postojećim ERP-om |
| Terenski otkupljivač / vagač | krajnji korisnik i izvor usvajanja ili otpora | ne | indirektno | minimalan broj koraka, offline rad, štampa, stabilnost |
| IT / eksterni tehnički savetnik | tehnički evaluator i risk gate | retko | da | arhitektura, bezbednost, monitoring, backup, podrška |
| Vlasnikov poverljivi savetnik / član porodice | skriveni influencer ili veto | retko | da | poverenje u dobavljača, reputacija, razumljiv poslovni slučaj |

`HYPOTHESIS`: U malim i srednjim hladnjačama jedna osoba često ima dve ili više navedenih uloga. Prodaja mora identifikovati funkcije, ne samo titule.

---

# 4. Profil A — vlasnik / generalni direktor

## 4.1 Poslovni kontekst

`FACT`: Ova uloga je krajnje odgovorna za sezonu, novac, mrežu stanica i reputaciju firme. Potreban joj je pregled bez ulaska u svaki dokument.

## 4.2 Desired progress

Želi da pređe iz stanja:

> „Informacije dobijam kroz pozive, poruke i ljude koji svako vidi samo svoj deo.“

u stanje:

> „Vidim gde nastaje odstupanje i mogu da reagujem pre nego što postane finansijski ili reputacioni problem.“

On ne kupuje PWA, dashboard ili audit log. Kupuje:

- predvidljiviju sezonu;
- ranije upozorenje;
- kontrolu bez stalnog zvanja ljudi;
- manju zavisnost od jedne ključne osobe;
- mogućnost rasta mreže bez proporcionalnog rasta haosa.

## 4.3 Dominantne motivacije

- **Kontrola:** želi jednu pouzdanu sliku poslovanja.
- **Zaštita kapitala:** želi da smanji greške, gubitke, dupliranja i nevidljive obaveze.
- **Brzina odluke:** želi podatke dok još može da utiče na ishod.
- **Skalabilnost:** želi rast bez administrativnog kolapsa.
- **Reputacija:** ne želi da nova tehnologija ugrozi sezonu, kooperante ili kupce.

## 4.4 Strahovi i lični rizik

`HYPOTHESIS`:

- da će implementacija poremetiti sezonu;
- da će zaposleni odbiti sistem i vratiti se na stare kanale;
- da će zavisiti od malog dobavljača ili jedne osobe;
- da će platiti širok sistem, a koristiti mali deo;
- da će promena otkriti ranije neuređene procese i odgovornosti;
- da će pred drugima izgledati kao osoba koja je donela pogrešnu odluku.

Najveći skriveni konkurent često nije drugi softver, nego odluka:

> „Izdržaćemo još ovu sezonu kao i do sada.“

## 4.5 Status quo nagrada

- ne mora sada da ulaže novac i vreme;
- ne preuzima javni rizik promene;
- postojeći ljudi već znaju improvizovani sistem;
- može privremeno da rešava probleme dodatnim pozivima i kontrolom;
- odlaže otvaranje pitanja vlasništva nad procesima i podacima.

## 4.6 Okidači za kupovinu

`HYPOTHESIS`, za validaciju:

- rast broja stanica ili kooperanata;
- odlazak ili preopterećenje ključnog administratora;
- ozbiljna greška u isplati, robi, ambalaži ili dokumentaciji;
- ulazak novog velikog kupca ili zahtev za boljom sledljivošću;
- nezadovoljstvo postojećim parcijalnim sistemom;
- potreba da firma funkcioniše bez stalnog ličnog mikromenadžmenta;
- priprema sledeće sezone i svest da postojeći način više ne skaluje.

## 4.7 Jezik koji otvara razgovor

Koristiti:

- „Kako danas dobijate pouzdanu sliku svih stanica dok je sezona u punom intenzitetu?“
- „Gde najčešće saznate za odstupanje kasnije nego što biste želeli?“
- „Koliko poslovanje trenutno zavisi od jedne ili dve osobe koje znaju gde su svi podaci?“
- „Šta biste morali da vidite da biste procenili da promena neće ugroziti sezonu?“
- „Ako se broj stanica poveća, koji deo sadašnjeg procesa prvi postaje usko grlo?“

Izbegavati:

- tehnički obilazak funkcija bez poslovnog konteksta;
- „digitalna transformacija“ kao praznu frazu;
- tvrdnje da je postojeći rad primitivan ili pogrešan;
- pritisak na odluku pre nego što je razjašnjen rizik implementacije;
- obećanje „potpune kontrole“ bez definisanih podataka i procesa.

## 4.8 Dokazni prag

Da bi prešao na sledeći korak, vlasnik obično mora da vidi:

1. da AgriX razume realan tok njegove firme;
2. da sistem radi u sezonskim, terenskim i offline uslovima;
3. da implementacija ima jasan obim, odgovornosti i plan povratka;
4. da postoji dokaz kod sličnog klijenta;
5. da poslovna korist nije samo „manje papira“, već bolja kontrola i manji rizik;
6. da dobavljač neće nestati posle prodaje.

## 4.9 Pozitivni buying signals

- uvodi druge odgovorne osobe u razgovor;
- daje konkretne brojeve stanica, korisnika, procesa i rokova;
- pita za implementaciju, migraciju, obuku i podršku;
- traži da demo prati njegov stvarni scenario;
- razmatra ko interno mora biti uključen;
- prihvata konkretan sledeći korak sa datumom.

## 4.10 Lažni signali

- opšta pohvala bez pristupa drugim stakeholderima;
- traženje cenovnika bez razgovora o obimu;
- „čujemo se posle sezone“ bez definisanog datuma i razloga;
- traženje velikog broja funkcija bez identifikovanog problema;
- zainteresovanost za tehnologiju, ali bez vlasnika poslovne posledice.

---

# 5. Profil B — operativni direktor / rukovodilac otkupa

## 5.1 Poslovni kontekst

Ova osoba živi posledice loše koordinacije. Ona je između vlasnika koji traži pregled i terena koji radi pod pritiskom.

## 5.2 Desired progress

> „Ne želim da ceo dan spajam informacije iz stanica, vozača, prijema i administracije. Želim da izuzeci budu vidljivi, a standardni tok da radi bez mog stalnog posredovanja.“

Kupuje:

- manje telefonske koordinacije;
- jasne statuse;
- brže rešavanje izuzetaka;
- disciplinu procesa bez ručnog nadzora svakog koraka;
- dokaz da je operacija izvršena kako je dogovoreno.

## 5.3 Motivacije

- operativni mir;
- manje kriznog rada;
- manje zavisnosti od usmenih dogovora;
- jasna odgovornost po koraku;
- mogućnost da rukovodi sistemom, ne pojedinačnim incidentima.

## 5.4 Strahovi

- da sistem neće pratiti realne izuzetke;
- da će biti dodatni sloj administracije;
- da će teren unositi nepotpune ili netačne podatke;
- da će on postati interna podrška za svaki problem;
- da će demo prikazati idealan tok, a ne haotičnu sezonu.

## 5.5 Jezik koji otvara razgovor

- „Koji deo dana Vam odlazi na prikupljanje statusa koji bi već trebalo da bude vidljiv?“
- „Gde se standardni tok najčešće pretvara u ručnu intervenciju?“
- „Koja tri izuzetka morate da vidite odmah, a ne na kraju dana?“
- „Kada se plan promeni, kako danas svi dobiju novu verziju informacije?“

## 5.6 Dokazni prag

- scenario od početka do kraja, uključujući izuzetak;
- real-time status, ali i audit istorija;
- jasna pravila odgovornosti;
- jednostavan rad za teren;
- dokaz da sistem smanjuje, a ne povećava broj koordinacija.

## 5.7 Buying signals

- opisuje konkretan incident iz prethodne sezone;
- crta ili objašnjava trenutni tok;
- identifikuje podatak koji mu nedostaje;
- predlaže pilot stanicu ili proces;
- traži uključivanje administratora i terenskog korisnika.

---

# 6. Profil C — centralni administrator / ključni Excel operator

## 6.1 Poslovni kontekst

`FACT`: Ova osoba često predstavlja operativno jezgro i jedina razume kako se podaci, dokumenti i izuzeci stvarno povezuju.

Ona može biti najbolji champion ili najjači tihi bloker.

## 6.2 Desired progress

> „Želim manje prekucavanja i manje traženja grešaka, ali ne želim da izgubim kontrolu, fleksibilnost i znanje koje mi danas omogućava da spasem situaciju.“

Kupuje:

- smanjenje ponovljenog rada;
- kontrolisane korekcije;
- proveru integriteta;
- manje odgovornosti za greške drugih;
- sistem koji poštuje realne dokumentne veze.

## 6.3 Dvostruka psihologija uloge

Ova uloga istovremeno želi rasterećenje i može se plašiti gubitka statusa.

`HYPOTHESIS`:

- postojeći ručni sistem joj daje stručni autoritet;
- automatizacija može izgledati kao umanjivanje njenog znanja;
- transparentan audit može povećati osećaj izloženosti;
- standardizacija može biti doživljena kao gubitak fleksibilnosti;
- neuspešna implementacija će prvo pasti na nju.

Zato prodaja ne sme da je tretira kao „osobu koju će sistem zameniti“. Ispravna pozicija je:

> AgriX uklanja prekucavanje i lov na greške, ali povećava vrednost njenog procesnog znanja.

## 6.4 Jezik koji otvara razgovor

- „Koje provere danas radite zato što ne možete da verujete da su ulazni podaci potpuni?“
- „Koji deo procesa samo Vi znate da završite kada nastane izuzetak?“
- „Koja korekcija je najrizičnija zato što utiče na više dokumenata?“
- „Šta sistem mora da sačuva da Vam ne bi oduzeo potrebnu fleksibilnost?“
- „Koju vrstu greške najčešće ispravljate, iako nije nastala kod Vas?“

## 6.5 Dokazni prag

- prikaz konkretnog dokumentnog lanca;
- storno, korekcija i audit na stvarnom primeru;
- upozorenja i integritet podataka;
- jasno vlasništvo nad master podacima;
- eksport i kontrola, bez zaključavanja u crnu kutiju;
- obuka i podrška u prvom ciklusu.

## 6.6 Buying signals

- počinje da navodi izuzetke i rubne slučajeve;
- pita kako se radi korekcija, ne samo unos;
- daje uzorke dokumenata ili šema;
- razlikuje obavezno od „lepo bi bilo“;
- želi test sa realnim podacima.

## 6.7 Red flags

- očekuje da novi sistem preslika svaku istorijsku improvizaciju;
- traži neograničene individualne izmene bez standardizacije;
- odbija uključivanje drugih korisnika;
- insistira na paralelnom radu bez kriterijuma završetka;
- skriva procesne detalje da bi zadržala kontrolu.

---

# 7. Profil D — finansije / knjigovodstvo

## 7.1 Desired progress

> „Želim da operativni podaci stignu tačno, potpuno i proverljivo, bez toga da finansije naknadno rekonstruišu šta se dogodilo.“

Kupuje:

- integritet dokumenata;
- manje neusaglašenih salda;
- jasnu vezu operacije, obaveze i plaćanja;
- kontrolisan SEF i bankarski tok;
- audit i trag korekcije;
- saradnju sa postojećim knjigovodstvenim sistemom.

## 7.2 Glavni strahovi

- da AgriX pokušava da zameni ERP bez dovoljne dubine;
- da će nastati još jedan izvor podataka koji treba usaglašavati;
- da će operativa unositi podatke bez računovodstvene kontrole;
- da će automatizacija proizvesti brže, ali pogrešne dokumente;
- da odgovornost za grešku neće biti jasna.

## 7.3 Jezik koji otvara razgovor

- „Na kom mestu operativni podatak danas postaje finansijski dokument?“
- „Koja neusaglašenost Vam najčešće stigne tek kada treba izvršiti isplatu ili zatvoriti period?“
- „Koji podatak se danas ponovo unosi u BizniSoft, PANTHEON ili drugi sistem?“
- „Šta mora biti provereno pre nego što dokument ili nalog može dalje?“

## 7.4 Dokazni prag

- jasan sistem-of-record model;
- integracija ili kontrolisan prenos ka postojećem ERP-u;
- audit i zabrana nevidljivog brisanja;
- pravila odobravanja i validacije;
- konkretni SEF, banka, saldo i isplatni scenariji.

---

# 8. Profil E — terenski otkupljivač / vagač

## 8.1 Desired progress

> „Želim da unesem tačan otkup i izdam dokument bez čekanja, komplikovanih menija i straha da će internet ili štampač zaustaviti red.“

Ova uloga uglavnom ne bira dobavljača, ali odlučuje da li će implementacija stvarno živeti.

## 8.2 Motivacije

- brzina;
- jednostavnost;
- mali broj odluka pod pritiskom;
- pouzdan rad bez stabilnog interneta;
- jasan status da li je podatak sačuvan i sinhronizovan;
- laka korekcija kroz odobren proces.

## 8.3 Strahovi

- da će sistem usporiti rad pred kooperantima;
- da će biti okrivljen za problem mreže ili uređaja;
- da će morati da pamti komplikovana pravila;
- da će izgubiti mogućnost da „brzo reši“ nestandardnu situaciju;
- da će ga sistem nadzirati bez podrške i jasnih pravila.

## 8.4 Jezik koji otvara razgovor

Ne pitati apstraktno šta želi od softvera. Posmatrati ili simulirati stvarni tok:

- „Pokažite mi najbrži standardni otkup.“
- „Šta se desi kada internet nestane usred unosa?“
- „Koja greška se najlakše napravi kada je red najveći?“
- „Kada štampa ne uspe, šta radite dalje?“
- „Koji korak biste prvi izbacili da možete?“

## 8.5 Dokazni prag

- praktičan test na uređaju;
- offline scenario;
- štampa i reprint;
- jasan sync status;
- minimalan broj polja i koraka;
- obuka zasnovana na zadatku, ne prezentaciji.

---

# 9. Profil F — IT / eksterni tehnički savetnik

## 9.1 Desired progress

> „Želim da poslovanje dobije funkcionalnost bez stvaranja nekontrolisanog tehničkog, bezbednosnog i support rizika.“

## 9.2 Motivacije

- jasna arhitektura i vlasništvo nad podacima;
- kontrola pristupa;
- backup, monitoring i oporavak;
- predvidiv deployment i update;
- granice podrške;
- dokumentovan integracioni model.

## 9.3 Strahovi

- skriveni single point of failure;
- oslanjanje na osnivača;
- nejasna bezbednost i prava pristupa;
- nekontrolisane lokalne izmene;
- nedokumentovane integracije;
- shadow IT koji kasnije IT mora da održava.

## 9.4 Jezik koji otvara razgovor

- „Koji su Vaši minimalni uslovi za sistem koji ulazi u kritični sezonski tok?“
- „Koji failure scenario moramo zajedno da prođemo pre odobrenja?“
- „Ko je vlasnik pristupa, backup-a, uređaja i eskalacije?“
- „Koji podaci moraju ostati interoperabilni sa postojećim sistemima?“

## 9.5 Dokazni prag

- arhitektonska dokumentacija;
- monitoring i update model;
- backup/recovery procedura;
- permissions model;
- support SLA i eskalacija;
- zabrana trajnih klijentskih forkova;
- dokaz rada u realnom okruženju.

---

# 10. Profil G — skriveni influencer / poverljivi savetnik

U porodičnim ili vlasnički koncentrisanim firmama formalna organizaciona šema često ne pokazuje stvarni centar poverenja.

`HYPOTHESIS`: Konačna odluka može zavisiti od člana porodice, eksternog knjigovođe, dugogodišnjeg saradnika ili konsultanta koji nema formalnu funkciju u projektu.

## 10.1 Rizik

Ova osoba može blokirati odluku rečenicom:

- „Nemoj sada pred sezonu.“
- „Ko zna koliko će ta firma trajati.“
- „Već imamo BizniSoft.“
- „Ljudi to neće koristiti.“
- „Bolje da ostane kako jeste.“

## 10.2 Pravilo

Do kraja discovery faze mora biti postavljeno pitanje:

> „Ko će, pored ljudi koji su danas uključeni, imati presudan uticaj na odluku ili može zaustaviti uvođenje?“

Ne pokušavati zaobići ovu osobu. Napraviti materijal koji glavnom kontaktu omogućava da interno objasni poslovni slučaj i rizik implementacije.

---

# 11. Psihološki obrasci preko svih persona

## 11.1 Loss aversion

Kupci snažnije osećaju rizik neuspešne promene nego potencijalnu dobit poboljšanja. Zato poruka mora istovremeno pokazati:

- cenu statusa quo;
- kontrolisan način promene;
- dokaz da rizik implementacije ima vlasnika i proceduru.

Ne koristiti zastrašivanje ili izmišljene gubitke.

## 11.2 Status quo bias

Postojeći sistem je poznat, čak i kada je neefikasan. Prodaja mora priznati šta u njemu radi dobro i pokazati gde više ne skaluje.

Pogrešno:

> „Papir i Excel su zastareli.“

Ispravno:

> „Excel Vam je verovatno dao fleksibilnost da izgradite proces. Pitanje je na kom obimu ta fleksibilnost počinje da stvara zavisnost od pojedinaca i ručne kontrole.“

## 11.3 Implementation anxiety

Kod sezonski kritičnog sistema, kupac ne procenjuje samo proizvod već i verovatnoću bezbednog uvođenja. Demo bez implementation narrative-a nije dovoljan.

## 11.4 Identity and competence protection

Ljudi se opiru promeni kada ona implicitno poručuje da su do sada radili loše. AgriX mora da poštuje postojeće procesno znanje i da ga pretvori u standard, a ne da ga omalovažava.

## 11.5 Decision diffusion

Što je više uloga, lakše je da niko ne preuzme sledeći korak. Svaki ozbiljan razgovor završava se konkretnim vlasnikom, aktivnošću i datumom.

---

# 12. Kako se profil koristi u prodaji

Pre svakog razgovora prodavac popunjava:

- primarna uloga sagovornika;
- dodatna funkcija koju verovatno obavlja;
- šta je njegov merljivi poslovni ishod;
- čega se lično i profesionalno izlaže promenom;
- koji dokaz je najverovatnije potreban;
- koju internu osobu mora uključiti;
- koja hipoteza mora biti proverena, a ne pretpostavljena.

Posle razgovora profil se ne upisuje kao etiketa tipa „konzervativan“ ili „težak“. Beleže se ponašanja i dokazi:

- „tražio rollback plan“;
- „odbija promenu tokom sezone“;
- „želi da administrator potvrdi korekcije“;
- „uveo vlasnika u sledeći sastanak“;
- „nije naveo posledicu problema“.

---

# 13. Validacioni intervju — pitanja za prvih 20 razgovora

Ne postavljaju se sva pitanja svakoj osobi. Biraju se prema ulozi i toku razgovora.

1. Šta je poslednji događaj zbog kog ste menjali način rada?
2. Koji problem tokom sezone saznate kasnije nego što biste želeli?
3. Ko danas spaja podatke kada se izvori ne slažu?
4. Šta bi moglo da krene loše pri uvođenju novog sistema?
5. Ko bi prvi osetio posledice neuspešne implementacije?
6. Ko bi mogao da zaustavi odluku, iako nije na sastanku?
7. Koji dokaz bi Vam bio dovoljan da podržite pilot?
8. Šta postojeći način rada radi dovoljno dobro da ne želite da ga izgubite?
9. Koji deo procesa mora ostati fleksibilan?
10. Kako biste za godinu dana znali da je odluka bila dobra?
11. Šta bi moralo da se dogodi da projekat postane prioritet ove godine?
12. Zašto taj problem još nije rešen?

---

# 14. Hipoteze koje moraju biti validirane

| ID | Hipoteza | Način validacije | Status |
|---|---|---|---|
| PSY-H01 | Vlasniku je rizik implementacije veća prepreka od same cene. | 10+ owner razgovora; poređenje prigovora i sledećih koraka | OPEN |
| PSY-H02 | Ključni administrator je najčešći tihi veto u malim firmama. | stakeholder map za 15 prilika | OPEN |
| PSY-H03 | „Imamo BizniSoft“ najčešće znači zaštitu statusa quo, ne punu funkcionalnu pokrivenost. | pitati koji procesi ostaju van sistema | OPEN |
| PSY-H04 | Najjači owner message je ranije otkrivanje odstupanja, ne ušteda administrativnih sati. | A/B poruke i kvalitativni odziv | OPEN |
| PSY-H05 | Terenski otpor je više vezan za brzinu i pouzdanost nego za opšti otpor tehnologiji. | posmatranje korisnika i pilot feedback | OPEN |
| PSY-H06 | Skriveni savetnik ili član porodice utiče na značajan deo konačnih odluka. | eksplicitna decision-map pitanja | OPEN |
| PSY-H07 | Reference sa merljivim rezultatom i javnim imenom zajedno imaju najveći uticaj. | beleženje proof asset-a koji pomera priliku | OPEN |

---

# 15. Zabranjene prakse

- izmišljanje straha ili pritiska;
- pripisivanje ličnih osobina na osnovu titule;
- manipulativno korišćenje porodičnih ili vlasničkih odnosa;
- stvaranje lažne hitnosti;
- skrivanje rizika implementacije;
- omalovažavanje postojećeg sistema ili zaposlenih;
- predstavljanje hipoteze kao potvrđene činjenice;
- obećavanje poslovnog rezultata bez definisane početne vrednosti i načina merenja.

---

# 16. Status Commercial Operating System-a

| Oblast | Status |
|---|---|
| Commercial Operating System | DONE v1 |
| Market Positioning | NOT STARTED |
| Psychological Profiles | DRAFT v1 — VALIDATION |
| Buying Process | NOT STARTED |
| Sales Process | NOT STARTED |
| Discovery Playbook | NOT STARTED |
| Email Sequences | NOT STARTED |
| Call Playbooks | NOT STARTED |
| Demo Playbook | NOT STARTED |
| Objection Handling | NOT STARTED |
| Negotiation | NOT STARTED |
| Case Studies | NOT STARTED |
| ROI Calculator | NOT STARTED |
| CRM Pipeline | NOT STARTED |
| KPI Dashboard | NOT STARTED |
| Annual Sales Calendar | NOT STARTED |
