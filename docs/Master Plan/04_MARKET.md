# 04 — Tržište

**Status:** Review  
**Vlasnik:** osnivač AgriX-a  
**Horizont:** 2026–2030  
**Poslednje ažuriranje:** 2026-07-23  
**Povezani dokumenti:** `02_STRATEGY.md`, `03_CUSTOMERS_AND_JOBS.md`, `05_COMPETITION.md`, `10_PRICING_AND_PACKAGING.md`, `14_GO_TO_MARKET.md`  
**Primarni dokazni skup:** `docs/Market Intelligence/01_Market_Data/APR/`

---

## 1. Svrha poglavlja

Ovo poglavlje definiše:

- koliki je dokazivi tržišni univerzum u Srbiji;
- koji deo tog univerzuma realno odgovara AgriX Enterprise proizvodu;
- koliki su TAM, SAM i realni SOM;
- kako ekonomska koncentracija utiče na redosled prodaje;
- kako se Enterprise tržište pretvara u distribucioni kanal za Gazdinstvo, GGAP i kiosk terminale;
- koje tržišne tvrdnje su potvrđene, a koje i dalje predstavljaju hipotezu;
- koje dodatne podatke treba prikupiti pre donošenja velikih investicionih odluka.

Ovo nije promotivni dokument. Broj registrovanih firmi nije isto što i broj kupaca, prihod firme nije isto što i operativna složenost, a strateški cilj nije isto što i prognoza.

---

## 2. Statusi tvrdnji i nivo pouzdanosti

U ovom poglavlju koriste se sledeći statusi:

- `FACT` — potvrđeno registrom, proizvodom, ugovorom ili direktnim poslovnim događajem;
- `MEASURED` — izmereno u APR analizi ili postojećem poslovanju AgriX-a;
- `INFERENCE` — zaključak izveden iz više potvrđenih podataka;
- `HYPOTHESIS` — razumna, ali još nepotvrđena tržišna pretpostavka;
- `TARGET` — poslovni cilj, ne prognoza;
- `DECISION` — usvojeni strateški princip.

Nivoi pouzdanosti:

- **Visok** — podatak je direktan, ponovljiv i jasno ograničen;
- **Srednji** — podatak je dobar, ali ima poznate strukturne nedostatke;
- **Nizak** — mali uzorak, ekstrapolacija ili nepotvrđena pretpostavka.

---

## 3. Definicija tržišta

AgriX ne cilja sve firme u poljoprivredi i prehrambenoj industriji.

Primarno tržište čine organizacije koje imaju dovoljno složen operativni tok da im je potreban povezan sistem za:

- kooperante, parcele i pripremu sezone;
- više otkupnih tačaka ili drugih operativnih lokacija;
- prijem robe i dokumentaciju na mestu nastanka;
- repromaterijal, agrohemiju, ambalažu ili druga zaduženja;
- transport, dispečiranje, prijem i lager;
- sledljivost;
- kupce, otpremu i fakturisanje;
- SEF, banku, salda i isplate;
- management kontrolu;
- GGAP ili sličan dokumentacioni pritisak.

`DECISION`: osnovna jedinica tržišta za Enterprise je **firma sa organizovanim otkupnim i povezanim operativnim tokom**, a ne svako pravno lice sa odgovarajućom šifrom delatnosti.

`DECISION`: AgriX se prodaje kao operativni sistem koji može da koegzistira sa BizniSoftom, PANTHEON-om ili drugim računovodstvenim ERP-om. Tržište se ne sužava samo na firme koje žele da zamene knjigovodstveni sistem.

---

## 4. Dokazna osnova

### 4.1 APR snapshot

APR pipeline je 22. jula 2026. obradio firme sa šiframa delatnosti:

- `1039` — ostala prerada i konzervisanje voća i povrća;
- `4631` — trgovina na veliko voćem i povrćem.

Rezultati:

| Pokazatelj | Vrednost | Status | Pouzdanost |
|---|---:|---|---|
| Ukupno pronađenih firmi u dve šifre | 1.644 | `MEASURED` | visok |
| Aktivne firme | 1.516 | `MEASURED` | visok |
| Firme u likvidaciji | 92 | `MEASURED` | visok |
| Firme u stečaju | 36 | `MEASURED` | visok |
| Aktivne firme sa podatkom o prihodu | 1.217 | `MEASURED` | visok |
| Pokrivenost aktivnog skupa prihodima | 80,3% | `MEASURED` | visok |
| Ukupni prihodi pokrivenog aktivnog skupa | 218.994.141.000 RSD | `MEASURED` | srednji–visok |
| Medijana prihoda | 23.027.000 RSD | `MEASURED` | srednji–visok |
| 90. percentil prihoda | 398.317.800 RSD | `MEASURED` | srednji–visok |

APR novčane vrednosti iz izvora normalizovane su iz hiljada dinara u pune RSD vrednosti. Pipeline čuva jedinicu u metadata fajlovima i odbija analizu ako finansijska jedinica nije potvrđena.

### 4.2 Koncentracija prihoda

| Grupa | Udeo prihoda analiziranog skupa | Status |
|---|---:|---|
| Top 10 firmi | 24,2% | `MEASURED` |
| Top 25 firmi | 39,4% | `MEASURED` |
| Top 50 firmi | 52,7% | `MEASURED` |
| Top 100 firmi | 66,8% | `MEASURED` |

`INFERENCE`: tržište je ekonomski koncentrisano. Mali broj velikih firmi nosi dominantan deo prihoda, dok postoji dugačak rep manjih subjekata.

`INFERENCE`: AgriX ne treba da rasporedi prodajni napor ravnomerno na svih 1.516 aktivnih firmi. Account-based pristup prema najrelevantnijih 100–300 firmi ima veći očekivani povrat od neselektivnog masovnog marketinga.

### 4.3 Interni podaci AgriX-a

- `FACT`: postoje tri aktivna klijenta;
- `FACT`: tipična postojeća firma ima približno deset stanica i oko 100 kooperanata;
- `FACT`: jedan centralni desktop korisnik i management PWA korisnici predstavljaju tipičnu trenutnu konfiguraciju;
- `FACT`: onboarding je do sada remote i traje približno jedan dan;
- `MEASURED`: trenutni support je približno jedan poziv nedeljno po postojećoj bazi, uz veliku razliku između trivijalnog pitanja i ozbiljnog buga;
- `FACT`: dva rana tržišna upita koristila su BizniSoft i SEF, ali su prvenstveno tražila bolju kontrolu ulaznog toka robe;
- `FACT`: jedan kupac za malinu ugovorio je AgriX za dve firme, sa potrebom za otkupnim listovima, prijemnicama, stanicama i reversom ambalaže;
- `FACT`: upit iz sektora duvana pokazuje da relevantno tržište postoji i van šifara `1039` i `4631`.

`INFERENCE`: APR skup je koristan donji temelj za tržište voća i povrća, ali nije kompletan AgriX tržišni univerzum.

---

## 5. Šta APR podaci dokazuju, a šta ne dokazuju

### APR podaci dokazuju

- da ranija neformalna procena od 500–1.000 relevantnih aktera nije dovoljna kao tržišna osnova;
- da samo dve relevantne šifre sadrže 1.516 aktivnih firmi;
- da tržište ima veliki broj malih firmi i ekonomski snažan gornji segment;
- da postoji dovoljno veliki account universe za višegodišnju B2B prodaju;
- da top 100 firmi zaslužuje poseban prodajni i istraživački tretman.

### APR podaci ne dokazuju

- da svih 1.516 firmi organizuje otkup preko mreže kooperanata;
- da sve imaju više stanica, sopstvenu logistiku ili potrebu za AgriX-om;
- da prihod određuje broj stanica, broj korisnika ili spremnost za kupovinu;
- da je firma sa visokim prihodom automatski bolji kupac od operativno složene firme sa manjim prihodom;
- da su sve relevantne firme registrovane pod šiframa `1039` ili `4631`;
- da finansijski endpoint predstavlja kompletnu višegodišnju istoriju;
- koliko firmi koristi Excel, papir, specijalizovani softver ili sopstveno rešenje;
- koliki je willingness-to-pay.

`DECISION`: prihod je signal za prioritet istraživanja i prodaje, ali ne sme biti jedini ICP kriterijum niti jedina osnova za cenu.

---

## 6. Tržišna struktura i segmenti

### 6.1 Tier A — strateški Enterprise računi

Tipični signali:

- više stanica i veliki broj kooperanata;
- sopstvena ili kompleksna logistika;
- repromaterijal, ambalaža i finansijska zaduženja;
- velik dokumentacioni volumen;
- više operativnih uloga;
- izvoz, sledljivost ili GGAP pritisak;
- potreba za management kontrolom;
- interni champion i owner implementacije.

Kvantitativni početni proxy:

- približno gornjih 10% firmi sa dostupnim prihodima;
- oko 122 firme iz trenutnog revenue-covered skupa;
- prag 90. percentila: približno 398,3 miliona RSD godišnjih prihoda.

`HYPOTHESIS`: većina najvrednijih prvih Enterprise računa nalazi se u ovom segmentu, ali operativna složenost mora biti ručno potvrđena.

### 6.2 Tier B — standardni višestanični kupci

Tipični signali:

- 5–15 stanica;
- približno 100–500 kooperanata;
- centralna administracija;
- Excel, telefon, Viber/WhatsApp i nepovezane evidencije;
- ponovni unos između terena, prijema, dokumenata i finansija;
- jasna potreba za kontrolom, ali niža kompleksnost od Tier A.

`HYPOTHESIS`: Tier B je najverovatnije jezgro ponovljive Standard ponude i može biti brojniji od Tier A.

### 6.3 Tier C — jednostavniji i manji kupci

Tipični signali:

- jedna ili mali broj tačaka;
- mali dokumentacioni obim;
- jednostavna logistika;
- mali broj kooperanata;
- dominantna potreba za dokumentima, a ne za kompletnim operativnim sistemom.

`INFERENCE`: ovaj segment povećava broj potencijalnih kupaca, ali može imati niži ARR, veću cenovnu osetljivost i lošiji odnos onboarding/support troška prema prihodu.

`DECISION`: Tier C ne sme da diktira roadmap Enterprise proizvoda niti da pretvori AgriX u jeftin generator otkupnih listova.

### 6.4 Adjacent segmenti

Potencijalni segmenti van trenutnog APR modela:

- duvan;
- žitarice i uljarice;
- mleko i stočarski otkup;
- lekovito bilje;
- druge kulture i organizatori proizvodnje;
- zadruge i proizvođačke grupe sa drugačijom šifrom delatnosti.

`FACT`: upit iz sektora duvana potvrđuje da adjacent demand postoji.

`HYPOTHESIS`: adjacent segmenti mogu značajno povećati TAM, ali svaki zahteva proveru procesa, dokumentacije, regulatornih razlika i potrebnih izmena proizvoda.

---

## 7. TAM — total addressable market

TAM se prikazuje u više slojeva jer jedan broj stvara lažnu preciznost.

### 7.1 Registry TAM — usko definisan tržišni univerzum

`MEASURED`: 1.516 aktivnih firmi u šiframa `1039` i `4631`.

Ovo je najčvršći trenutno dostupan broj, ali predstavlja **registracioni univerzum**, ne broj spremnih kupaca.

### 7.2 Commercial TAM — sve firme kojima bi AgriX mogao da reši relevantan problem

Commercial TAM uključuje:

- relevantne firme iz APR skupa;
- firme iz drugih šifara koje organizuju otkup;
- zadruge i proizvođačke grupe;
- adjacent kulture i sektore.

`HYPOTHESIS`: commercial TAM je veći od 1.516 firmi, ali trenutno nema dovoljno dokaza za pouzdan konačan broj.

`DECISION`: u eksternoj komunikaciji ne tvrditi da AgriX ima „tržište od 1.516 kupaca“. Ispravno je tvrditi da APR identifikuje 1.516 aktivnih firmi u dve relevantne šifre i da se stvarni product-fit deo tek kvalifikuje.

### 7.3 Revenue TAM

Revenue TAM ne treba zaključati pre odobrenja `10_PRICING_AND_PACKAGING.md` i stvarnih willingness-to-pay podataka.

`DECISION`: ne množiti svih 1.516 firmi jednom cenom i predstavljati rezultat kao realan prihodovni TAM.

---

## 8. SAM — serviceable available market

SAM je deo tržišta koji trenutni ili kratkoročno planirani AgriX može kvalitetno da onboarduje i podrži.

### 8.1 Donja granica

Gornjih 10% firmi sa dostupnim prihodima daje približno 122 računa.

Ovo je konzervativan kvantitativni proxy za velike firme, ali isključuje:

- operativno kompleksne firme ispod revenue praga;
- firme bez dostupnih finansijskih podataka;
- relevantne adjacent delatnosti.

### 8.2 Planerski SAM za Srbiju

`HYPOTHESIS`: realni planerski SAM za AgriX Enterprise u Srbiji trenutno iznosi približno **150–300 firmi**.

Obrazloženje:

- donju osnovu čini oko 122 firme iz gornjeg revenue decila;
- dodaju se višestanične i operativno kompleksne firme ispod tog praga;
- oduzimaju se firme koje ne organizuju otkup, nemaju dovoljan procesni problem ili nemaju kupovnu spremnost;
- adjacent segmenti nisu potpuno uračunati.

Pouzdanost ove procene je **niska–srednja** dok se ne uradi ručna klasifikacija najmanje top 300 APR računa.

### 8.3 Uslovi da firma uđe u SAM

Firma mora ispuniti većinu sledećih kriterijuma:

1. organizovan otkup ili organizovana proizvodnja;
2. više operativnih tačaka, uloga ili dokumentnih tokova;
3. dovoljan sezonski intenzitet i rizik greške;
4. potreba za centralnom kontrolom;
5. spremnost da prihvati standardan proizvod i konfiguraciju;
6. champion i owner implementacije;
7. minimalna tehnička i organizaciona spremnost;
8. ekonomska opravdanost onboardinga i supporta.

---

## 9. SOM — serviceable obtainable market

SOM mora da odražava prodajni kapacitet, onboarding, support i reputacioni rizik, ne samo veličinu tržišta.

### 9.1 Narednih 12–18 meseci

Postojeća baza: 3 firme.

| Scenario | Ukupno aktivnih Enterprise firmi | Status |
|---|---:|---|
| Konzervativni | 6–8 | `HYPOTHESIS` |
| Bazni | 8–12 | `HYPOTHESIS` |
| Ubrzani uz potvrđen readiness | 12–15 | `HYPOTHESIS` |

Ovo nije unapred postavljen limit prodaje. Broj novih klijenata treba da raste sa readiness score-om, standardizacijom onboardinga i support kapacitetom.

### 9.2 Horizont 3–4 godine u Srbiji

| Scenario | Aktivne Enterprise firme | Udeo registry TAM-a | Udeo planerskog SAM-a |
|---|---:|---:|---:|
| Konzervativni | 20 | 1,3% | 6,7–13,3% |
| Bazni | 40–60 | 2,6–4,0% | 13,3–40,0% |
| Stretch | 100 | 6,6% | 33,3–66,7% |
| Strateški cilj | 200 | 13,2% | 66,7–133,3% |

`TARGET`: najmanje 200 firmi u naredne 3–4 godine ostaje strateška ambicija iz `02_STRATEGY.md`.

`INFERENCE`: na osnovu trenutnog planerskog SAM-a, 200 firmi nije bazna prognoza za Srbiju. Za njegovo dostizanje verovatno su potrebni:

- širi adjacent segmenti;
- regionalno širenje;
- znatno veći sales kapacitet;
- ponovljiv onboarding koji ne zavisi od osnivača;
- standardizovana podrška;
- jasniji low-touch paket za deo tržišta;
- dokazano zadržavanje i preporuke.

`DECISION`: finansijski model ne sme koristiti 200 firmi kao bazni scenario bez posebnog prikaza potrebnih kapaciteta i verovatnoće.

---

## 10. Tržište AgriX Enterprise proizvoda

### 10.1 Glavni tržišni problem

Relevantne firme često ne traže zamenu knjigovodstvenog ERP-a. Traže kontrolu operativnog toka koji počinje pre knjigovodstva:

- ko je predao robu;
- na kojoj stanici;
- po kojoj ceni, klasi i količini;
- koja dokumentacija je nastala;
- šta je preuzeto i transportovano;
- šta je primljeno, uskladišteno i otpremljeno;
- šta treba fakturisati, naplatiti ili isplatiti;
- gde postoji odstupanje ili neusaglašenost.

Rani upiti potvrđuju da BizniSoft i SEF mogu ostati u firmi, dok AgriX rešava ulazni, terenski i operativni tok.

### 10.2 Tržišna spremnost

`INFERENCE`: tržište je spremnije za specijalizovani operativni sloj nego za poruku „zamenite ceo ERP“.

`INFERENCE`: kupac će procenjivati AgriX kroz četiri pitanja:

1. da li sistem radi tokom sezone i bez stabilnog interneta;
2. ko kontroliše i može da izveze podatke;
3. koliko traje implementacija i obuka;
4. šta se dešava ako nastane bug, kvar ili prekid.

`DECISION`: trust, kontinuitet sezone, backup/export i dokazani onboarding moraju biti deo tržišne ponude, a ne tehnički dodatak.

---

## 11. Tržište kiosk terminala i termalne štampe

Kiosk nije samostalno tržište; on je hardversko-operativni sloj Enterprise proizvoda.

Interna polazna činjenica:

- tipična postojeća firma ima približno deset stanica.

Scenario broja terminala:

| Broj Enterprise firmi | Približan broj stanica/terminala | Status |
|---:|---:|---|
| 10 | 100 | `INFERENCE` |
| 50 | 500 | `INFERENCE` |
| 100 | 1.000 | `INFERENCE` |
| 200 | 2.000 | `INFERENCE` |

Ovo nisu jedinstveni tržišni podaci, već ekstrapolacija trenutnog proseka.

`DECISION`: hardverska nabavka, zaliha i support ne smeju se planirati prema ukupnom TAM-u, već prema potpisanim firmama, broju potvrđenih stanica i sezonskom rollout planu.

`INFERENCE`: kiosk može povećati switching cost i dubinu korišćenja AgriX-a, ali može i pretvoriti softverski problem u terenski support problem ako uređaji, štampa i daljinsko upravljanje nisu standardizovani.

---

## 12. Tržište AgriX Gazdinstva

Gazdinstvo ima dva tržišna puta:

1. B2B2C distribucija preko Enterprise klijenata;
2. direktna prodaja gazdinstvima.

### 12.1 B2B2C distribucioni universe

Interna polazna činjenica:

- tipična postojeća Enterprise firma ima približno 100 kooperanata.

Scenario kooperantskih odnosa:

| Enterprise firme | Kooperantski odnosi | Plaćeni nalozi pri 5% konverzije |
|---:|---:|---:|
| 10 | oko 1.000 | oko 50 |
| 50 | oko 5.000 | oko 250 |
| 100 | oko 10.000 | oko 500 |
| 200 | oko 20.000 | oko 1.000 |

Statusi:

- 100 kooperanata po firmi: `FACT` za tipičnu trenutnu konfiguraciju, ali mali uzorak;
- 5% konverzije: `HYPOTHESIS`;
- rezultati tabele: `INFERENCE` niske pouzdanosti.

Ovi brojevi predstavljaju odnose firma–kooperant, ne nužno jedinstvena gazdinstva. Jedan proizvođač može sarađivati sa više firmi.

### 12.2 Standalone tržište Gazdinstva

`HYPOTHESIS`: standalone tržište može biti veće od Enterprise-distribuiranog tržišta, ali trenutno nema dovoljno dokaza o:

- broju digitalno spremnih voćara;
- aktivaciji nakon registracije;
- sezonskoj i godišnjoj retenciji;
- spremnosti da plate 19 EUR ili 39 EUR;
- support trošku po gazdinstvu;
- kanalu akvizicije sa prihvatljivim CAC-om.

`DECISION`: Gazdinstvo ne sme biti vrednovano kao veliki samostalni prihod samo na osnovu broja poljoprivrednih gazdinstava u Srbiji. Prvo se dokazuju aktivacija, retencija i plaćena konverzija kroz Enterprise mrežu.

---

## 13. Tržište AgriX GGAP proizvoda

GGAP tržište nije jednako broju svih APR firmi niti broju svih kooperanata.

Najbolji početni kupci su:

- izvoznici;
- firme i grupe proizvođača sa većim brojem gazdinstava;
- organizacije koje već imaju GGAP/quality koordinatora;
- firme koje ručno održavaju veliki broj dokaza, procedura i rokova;
- kupci kojima sertifikacija ili kupac nameću dokumentacionu disciplinu.

`INFERENCE`: GGAP fit će verovatno biti veći u Tier A i gornjem delu Tier B segmenta.

`HYPOTHESIS`: GGAP može povećati ARR po firmi, retenciju i vrednost Gazdinstva, ali tržišna veličina još nije kvantifikovana.

Za evidence-based procenu potrebni su:

- broj sertifikovanih firmi i proizvođačkih grupa;
- broj obuhvaćenih gazdinstava;
- broj konsultantskih i sertifikacionih projekata godišnje;
- trošak postojećeg ručnog održavanja dokumentacije;
- willingness-to-pay za kontinuirani readiness, a ne samo jednokratnu pripremu.

---

## 14. Sezonalnost

`FACT`: otkup je sezonski intenzivan i kritični periodi zavise od kulture.

Tržišne posledice:

- prodaja i ugovaranje moraju početi dovoljno rano pre sezone;
- onboarding ne sme biti planiran u trenutku kada je klijent već u punom otkupu;
- veliki release-i tokom sezone nose disproporcionalan reputacioni rizik;
- support potreba je neravnomerna;
- kupac može odlagati odluku van sezone, a zatim tražiti hitnu implementaciju pred početak otkupa;
- hardware lead time postaje deo prodajnog ciklusa;
- dokaz iz jedne uspešne sezone ima veću tržišnu vrednost od generičkog demo snimka.

`DECISION`: svaka prodajna prilika mora imati kulturu, očekivani početak sezone i poslednji bezbedan datum za onboarding.

`DECISION`: GTM KPI ne meri samo broj leadova, već i koliko je kvalifikovanih firmi ugovoreno dovoljno rano za bezbedan rollout.

---

## 15. Regionalna koncentracija u Srbiji

APR workbook sadrži analizu po opštinama, ali sama registraciona lokacija firme nije dovoljna za zaključak o stvarnoj mreži stanica, kooperanata i parcela.

`HYPOTHESIS`: prvi prodajni fokus treba da prati regione sa kombinacijom:

- velikog broja relevantnih firmi;
- visoke koncentracije voćarske i povrtarske proizvodnje;
- više otkupnih tačaka;
- postojećih referenci i preporuka;
- logističke dostupnosti;
- izraženog izvoznog i GGAP pritiska.

Trenutni kvalitativni fokus je Centralna i Zapadna Srbija, ali ovaj redosled još nije potvrđen kao konačna tržišna odluka.

`DECISION`: pre zaključavanja regionalnog sales plana spojiti:

1. APR firmu i opštinu;
2. promet i broj zaposlenih;
3. poznate kulture;
4. broj stanica i kooperanata iz discovery razgovora;
5. izvoz/GGAP signal;
6. postojeće reference i geografski referral efekat.

`TARGET`: napraviti listu prvih 100 računa sa regionalnim klasterima i ownerom sledeće prodajne akcije.

---

## 16. Regionalno širenje van Srbije

Planirana tržišta:

1. Bosna i Hercegovina;
2. Crna Gora;
3. Severna Makedonija;
4. Hrvatska;
5. druga relevantna tržišta nakon validacije.

Trenutno ne postoji evidence-based broj potencijalnih firmi za ova tržišta.

`DECISION`: regionalni TAM se ne dodaje srpskom TAM-u bez posebnog dataset-a i regulatornog pregleda.

Svaka zemlja mora imati:

- registar firmi i relevantne šifre;
- mapu dokumenata i poreskih pravila;
- e-fakture i bankarske integracije;
- lokalne obrasce otkupa;
- jezik i terminologiju;
- standarde zaštite podataka;
- lokalne reference ili partnera;
- procenu prodajnog i support troška.

`INFERENCE`: BiH i Crna Gora mogu biti operativno bliže, dok Hrvatska može imati viši potencijal i viši regulatorni/integracioni zahtev. Ovo ostaje hipoteza do zasebnog istraživanja.

---

## 17. Tržišne implikacije za prodaju

### 17.1 Prioritet nije „sve firme“

`DECISION`: prvi sales universe deli se na:

- **Tier A:** približno 100–150 strateških računa;
- **Tier B:** sledećih 150–250 standardnih potencijalnih računa;
- **Tier C:** dugačak rep koji se obrađuje jeftinijim kanalima i tek nakon standardizacije ponude.

Brojevi su planerski i moraju se potvrditi ručnom klasifikacijom.

### 17.2 Prodajni signal mora biti procesni

Obavezni signali:

- broj stanica;
- broj kooperanata;
- kulture i sezona;
- obim dokumentacije;
- repromaterijal i ambalaža;
- logistika i broj vozila;
- postojeći ERP;
- način prenosa podataka sa terena;
- SEF i banka;
- GGAP/izvoz;
- owner implementacije;
- poslednji bezbedan onboarding datum.

### 17.3 Sales-led, ne mass-self-service Enterprise

`INFERENCE`: Enterprise tržište zahteva poverenje, discovery, demo, validaciju procesa i kontrolisan onboarding.

`DECISION`: SEO i oglasi stvaraju i hvataju intent, ali kvalifikacija i zatvaranje ostaju sales-led dok se ne dokaže standardan low-touch paket.

### 17.4 Reference imaju klastersku vrednost

`INFERENCE`: jedna uspešna hladnjača u regionu i kulturi može otvoriti više sličnih računa kroz preporuke, zaposlene, knjigovođe, kooperante i poslovne veze.

`DECISION`: case study mora biti organizovan po procesu i kulturi, ne samo kao generičko svedočenje da je klijent zadovoljan.

---

## 18. Tržišne implikacije za proizvod

1. Enterprise mora da ostane operativni sistem, ne da se suzi na dokumente.
2. Integracije sa postojećim ERP-om imaju veću tržišnu vrednost od pokušaja da AgriX postane univerzalno knjigovodstvo.
3. Offline rad, backup, export i monitoring predstavljaju komercijalne trust funkcije.
4. PWA Otkupac i pouzdana štampa mogu proširiti upotrebu na stanicama, ali zahtevaju terensku pouzdanost.
5. Dispečer, Vozač, ambalaža, repromaterijal i sledljivost povećavaju fit za veće račune.
6. Gazdinstvo i GGAP treba da koriste Enterprise mrežu kao primarni kanal početne distribucije.
7. Klijentski forkovi bi uništili mogućnost da se SAM servisira; konfiguracija i zajednički kod ostaju strateški zahtev.

---

## 19. Ključni rizici tržišne procene

### 19.1 False-positive registry TAM

Firme mogu imati relevantnu šifru, ali ne i proces za koji je AgriX napravljen.

**Mitigacija:** ručno kvalifikovati top 300 računa i beležiti razlog fit/no-fit.

### 19.2 False-negative TAM

Relevantne firme mogu biti registrovane pod drugim šiframa.

**Mitigacija:** dodavati adjacent šifre tek nakon procesne validacije.

### 19.3 Revenue bias

Velika firma nije nužno složeniji kupac; mali prihod ne znači mali broj stanica ili nizak procesni bol.

**Mitigacija:** ICP score mora kombinovati prihod i operativne signale.

### 19.4 Mali interni uzorak

Tri klijenta nisu dovoljan uzorak za sigurne tržišne proseke.

**Mitigacija:** za svaki lead čuvati strukturisane podatke o stanicama, kooperantima, kulturama, dokumentima, ERP-u, sezoni i razlozima kupovine/odbijanja.

### 19.5 Cilj od 200 firmi postaje lažna prognoza

**Mitigacija:** u svim finansijskim modelima odvojiti `TARGET` 200 od konzervativnog, baznog i stretch SOM scenarija.

### 19.6 Sezonski reputacioni rizik

Jedan ozbiljan problem u sezoni može usporiti prodaju čitavog regionalnog klastera.

**Mitigacija:** readiness gating, staged rollout, monitoring, rezervna oprema i jasna incident komunikacija.

---

## 20. KPI-jevi tržišne validacije

### Kvartalno

- broj klasifikovanih APR računa;
- procenat top 300 računa sa poznatim brojem stanica;
- procenat sa poznatim brojem kooperanata;
- procenat sa poznatim ERP-om;
- procenat sa GGAP/izvoznim signalom;
- broj discovery razgovora;
- broj kvalifikovanih Tier A i Tier B prilika;
- demo-to-pilot i pilot-to-paid konverzija;
- win/loss razlozi;
- prosečno vreme od prvog kontakta do odluke;
- broj regionalnih referral leadova.

### Godišnje

- broj aktivnih Enterprise firmi;
- udeo u registry TAM-u;
- udeo u potvrđenom SAM-u;
- broj aktivnih stanica;
- broj Enterprise-distribuiranih Gazdinstvo naloga;
- plaćena konverzija Gazdinstva;
- broj GGAP kupaca i obuhvaćenih gazdinstava;
- ARR po segmentu;
- churn i razlog odlaska;
- support sati po firmi i stanici.

---

## 21. Obavezni istraživački backlog

### Prioritet 1 — top 300 računa

Za svaku firmu prikupiti:

- sajt i kontakt;
- kulture;
- broj stanica;
- procenu broja kooperanata;
- postojeći ERP/softver;
- repromaterijal i ambalažu;
- logistiku;
- izvoz i GGAP;
- ownera i championa;
- ICP score;
- sledeću prodajnu akciju.

### Prioritet 2 — tržište van dve APR šifre

- duvan;
- žitarice;
- zadruge;
- proizvođačke grupe;
- drugi relevantni sektori.

### Prioritet 3 — willingness-to-pay

Najmanje 15–20 strukturisanih razgovora po ključnom segmentu, uz testiranje:

- godišnje pretplate;
- implementacije;
- cene po stanici;
- dodatnih modula;
- kioska i hardvera;
- Gazdinstva;
- GGAP-a.

### Prioritet 4 — regionalna mapa

Spojiti APR, poljoprivrednu proizvodnju, postojeće reference i prodajni pipeline u jedinstvenu regionalnu prioritizaciju.

---

## 22. Strateške odluke iz ovog poglavlja

1. `DECISION`: 1.516 aktivnih firmi predstavlja registry TAM za dve šifre, ne potvrđeni broj kupaca.
2. `DECISION`: planerski Enterprise SAM za Srbiju privremeno se vodi kao 150–300 firmi.
3. `DECISION`: cilj od 200 firmi ostaje strateški stretch i ne koristi se kao bazna prognoza.
4. `DECISION`: prvi prodajni fokus je ručno kvalifikovanih top 100–300 računa.
5. `DECISION`: prihod je samo jedan deo ICP score-a.
6. `DECISION`: Enterprise ostaje sales-led i integracioni, a ne zamena za računovodstveni ERP.
7. `DECISION`: Gazdinstvo se prvo validira kroz Enterprise distribuciju.
8. `DECISION`: GGAP TAM se ne izmišlja bez posebnog dokaznog skupa.
9. `DECISION`: regionalni TAM se ne dodaje Srbiji bez zemlje-po-zemlje istraživanja.
10. `DECISION`: svaka tržišna procena se ažurira iz reproduktivnog pipeline-a i strukturisanog CRM istraživanja.

---

## 23. Zaključak

`FACT`: u samo dve relevantne APR šifre postoji 1.516 aktivnih firmi, što potvrđuje da je srpsko tržište dovoljno veliko za ozbiljan vertikalni B2B proizvod.

`FACT`: tržište je ekonomski koncentrisano; top 100 firmi nosi 66,8% prihoda analiziranog skupa.

`INFERENCE`: najveća kratkoročna prilika nije masovna prodaja svim firmama, već duboka kvalifikacija i osvajanje ograničenog broja visokofit računa.

`HYPOTHESIS`: realni trenutni Enterprise SAM u Srbiji je 150–300 firmi.

`INFERENCE`: 40–60 aktivnih Enterprise kupaca u horizontu 3–4 godine predstavlja zahtevnu, ali operativno verovatniju baznu zonu od 200 firmi. Sto firmi je stretch za Srbiju; 200 zahteva šire segmente, region i organizaciju koja više ne zavisi od osnivača.

AgriX ima tržište. Sledeći problem nije dokazati da firme postoje, već dokazati koje od njih imaju dovoljno jak problem, dovoljno visok fit, dovoljno poverenja i dovoljno kupovne spremnosti da postanu profitabilni i dugoročni klijenti.

---

## 24. Izvori i provenance

### Primarni izvori

- `docs/Market Intelligence/01_Market_Data/APR/Raw Data/apr_companies_1039_4631_2026-07-22.xlsx`
- `docs/Market Intelligence/01_Market_Data/APR/Clean Data/apr_companies_financials_1039_4631_2026-07-22.xlsx`
- `docs/Market Intelligence/01_Market_Data/APR/Processed/apr_market_validated_1039_4631_2026-07-22.xlsx`
- `docs/Market Intelligence/01_Market_Data/APR/Reports/apr_market_analysis_1039_4631_2026-07-22.xlsx`
- `docs/Market Intelligence/01_Market_Data/APR/Reports/apr_market_summary_1039_4631_2026-07-22.md`
- `docs/Market Intelligence/01_Market_Data/APR/Reports/apr_market_analysis_1039_4631_2026-07-22.metadata.json`

### Reproducibilnost

- APR pipeline se nalazi u `docs/Market Intelligence/01_Market_Data/APR/Scripts/`;
- finansijske vrednosti su normalizovane iz hiljada RSD u pune RSD;
- statusi se klasifikuju uz podršku za srpsku ćirilicu i latinicu;
- tržišna analiza se prekida ako je statusni scope prazan ili finansijska jedinica nije potvrđena;
- izveštaj na kojem se zasniva ovo poglavlje generisan je 22. jula 2026;
- relevantna RSD korekcija i regenerisani izlazi nalaze se na `main` branchu, commit `bdf1c6a36295249095b6c8d3afc419500cb33ef7`.

### Poznata ograničenja izvora

- APR šifre nisu savršena product-fit klasifikacija;
- 19,7% aktivnih firmi nema upotrebljiv prihod u trenutnom preseku;
- financial-statements endpoint se trenutno tretira kao jedan zapis po firmi;
- interni proseci stanica i kooperanata zasnovani su na maloj bazi klijenata;
- SAM, SOM, Gazdinstvo konverzija i regionalni redosled moraju se ažurirati novim dokazima.
