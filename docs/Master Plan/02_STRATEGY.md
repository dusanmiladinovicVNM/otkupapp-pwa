# 02 — Strategija i identitet AgriX-a

**Status:** Review  
**Vlasnik:** osnivač AgriX-a  
**Horizont:** 2026–2030  
**Poslednje ažuriranje:** 2026-07-27  
**Povezani dokumenti:** `02A_GGAP_STRATEGY.md`, `03_CUSTOMERS_AND_JOBS.md`, `07_PRODUCT_PORTFOLIO.md`, `09_QA_DECISION_LOG.md`, `DECISION_LOG.md`

---

## 1. Strateška polazna tačka

AgriX je vertikalni poslovni operativni sistem za organizovani otkup poljoprivrednih proizvoda, upravljanje gazdinstvom i kompletan dokumentacioni tok. Ne pokriva samo unos otkupa, već povezuje kooperante, parcele, repromaterijal, terenski rad, prijem robe, logistiku, transport, lager, prodaju, dokumentaciju, finansije, regulatorne obaveze, management kontrolu i farm-management funkcije.

Potvrđene činjenice, ciljevi i hipoteze:

- `FACT`: postoje tri aktivna klijenta;
- `FACT`: tipična firma ima približno deset otkupnih stanica, jednog centralnog desktop operatera, management PWA korisnike i oko 100 kooperanata;
- `FACT`: onboarding je do sada rađen remote i trenutno realno traje približno jedan dan;
- `TARGET`: standardizovati onboarding na približno pola dana za standardnog klijenta;
- `MEASURED`: dosadašnji support je približno jedan poziv nedeljno, od kratkih pitanja do višesatnih bug eskalacija;
- `FACT`: self-update, monitoring i zajednički codebase sa konfiguracijom već postoje;
- `FACT`: trajni klijentski forkovi nisu dozvoljeni;
- `FACT`: AgriX pokriva Dispečer i Vozač tok, SEF, bankarske izvode, rasknjižavanje, salda i pripremu naloga za plaćanje;
- `FACT`: management rola omogućava QR identifikaciju kooperanta, izbor parcela, preporuku količine agrohemije, barkod skeniranje, korpu i otpremnicu;
- `FACT`: Gazdinstvo pokriva karticu prema hladnjači, GIS, prognozu i upozorenja po parceli, pametno doziranje, tretmane, opremu, troškove i sezonski bilans;
- `TARGET`: GGAP modul treba da pokrije GGAP liste i kompletan dokumentacioni tok;
- `HYPOTHESIS`: u Srbiji postoji približno 500–1.000 relevantnih firmi; procena mora biti potvrđena u `04_MARKET.md`;
- `TARGET`: osvojiti najmanje 200 firmi u naredne 3–4 godine;
- `DECISION`: rast se planira prema fiksnom ciljnom broju klijenata, ne prema readiness cap-u (odluka 403; povlači STR-001).

AgriX više nije prototip, ali još nije dokazano skaliran proizvod. Sledeća faza mora dokazati da se širina sistema može pretvoriti u ponovljiv onboarding, pouzdanu podršku, održiv pricing i brz tržišni rast.

---

## 2. Identitet kompanije i tri proizvoda

AgriX je jedna platforma sa tri povezana proizvodna stuba — **Enterprise**, **Gazdinstvo** i **Savetnik** (odluke 269 i 401).

GGAP **nije** stub. GGAP je modul u okviru Enterprise-a i koriste ga isključivo hladnjače koje su već Enterprise klijenti; aktivacija modula otključava dodatne funkcije u Gazdinstvu (odluka 402, PRT4). Ranija odluka STR-012, koja je GGAP tretirala kao treći proizvodni stub, time je ukinuta.

### 2.1 AgriX Enterprise

Kompletan operativni sistem za hladnjače i druge firme koje organizuju otkup kroz mrežu stanica, kooperanata, vozila i kupaca.

Glavni domeni:

1. kooperanti, parcele, stanice, korisnici i priprema sezone;
2. repromaterijal, agrohemija, ambalaža i zaduženja;
3. terenski otkup, dokumenti, kiosk i termalna štampa;
4. Dispečer, vozači, kamioni, rute i statusi transporta;
5. prijem, lager, palete, prerada i sledljivost;
6. kupci, korpe, otpremnice, prodaja i fakture;
7. SEF i regulatorni tok;
8. banka, avansi, salda, rasknjižavanje i nalozi za plaćanje;
9. management, izveštaji, audit, storno i monitoring.

Enterprise treba da bude primarni operativni sistem firme čak i kada knjigovodstvo ostaje u BizniSoftu, PANTHEON-u ili drugom ERP-u.

#### Moduli uz Enterprise

Enterprise se prodaje kroz pakete Desktop i Mobile, uz posebno naplative module: Hladnjača/Proizvodnja, SEF, Banka, Dispatch (samo uz Mobile) i **GGAP**.

GGAP kao modul znači:

- kupac GGAP-a je uvek postojeći Enterprise klijent — GGAP se ne prodaje samostalno i nema sopstveni ICP;
- GGAP nema sopstveni stub-level product strategy, packaging ni unit economics; vodi se kao modul u Enterprise ekonomici;
- aktivacija modula otvara dodatne GGAP funkcije u Gazdinstvu kooperanata te hladnjače;
- do validacije GGAP ostaje **van komercijalne ponude** i nudi se samo kroz kontrolisan pilot (odluka 405).

Detaljan funkcionalni obuhvat GGAP-a ostaje u `02A_GGAP_STRATEGY.md`, koji se od odluke 402 čita kao strategija modula, ne stuba.

### 2.2 AgriX Gazdinstvo

Pun farm-management proizvod za kooperanta, a ne portal sa nekoliko read-only pregleda.

Glavni domeni:

1. kartica prema hladnjači: zaduženje, razduženje i saldo;
2. parcele, kulture, površine, GIS poligoni i satelitska mapa;
3. realna prognoza, rizici i upozorenja po konkretnoj parceli;
4. Digitalni agronom i pametno doziranje;
5. tretmani, karenca, oprema, vreme rada, lokacija i meteo snapshot;
6. knjiga polja i proizvodnja;
7. lager agrohemije;
8. troškovi po kategorijama i parcelama;
9. sezonski bilans ukupno i po parceli;
10. offline-first rad i sinhronizacija.

Gazdinstvo trenutno nije osnovni izvor prihoda, ali može postati glavni proizvod ili glavni prihod ako se potvrde aktivacija, retencija, willingness-to-pay i održiv support cost.

### 2.3 AgriX Savetnik

Treći ravnopravan stub (odluke 269, 401, PRT3). Proizvod za agronome i savetodavne službe — upravljački sloj nad većim brojem gazdinstava.

Glavni domeni:

1. istovremeni pregled portfelja gazdinstava;
2. radni nalozi i preporuke ka gazdinstvima;
3. praćenje izvršenja naloga i kontrola rada;
4. agronomska istorija po gazdinstvu, parceli i kulturi;
5. veza ka Gazdinstvo podacima uz saglasnost nosioca.

Komercijalni model je dvostruk: savetnik plaća alat — osnovica 150 € za do 10 gazdinstava, svako preko toga 15 € (odluka 341) — a gazdinstva u njegovom portfelju zadržavaju sopstvenu Pro pretplatu po kanalskoj ceni (odluka 340). Savetnik ne dobija proviziju za naloge u portfelju — podsticaj je sam alat, koji bez Pro naloga ne funkcioniše (odluka 345). Cena se objavljuje (odluka 347) i nalazi se u `docs/Sales/AgriX_Cenovnik_2027.pdf`.

Savetnik ima samostalnu registraciju i ne može javno na tržište bez sopstvene politike privatnosti; uloga rukovaoca za tok T13 ostaje otvorena do razrešenja LEG1 (`docs/Legal/AgriX_Mapa_tokova_podataka.pdf`).

### 2.4 Zajednički data flywheel

Tri stuba ne smeju postati tri silosa:

- Enterprise proizvodi transakcije, dokumente, finansijske podatke i sledljivost;
- Gazdinstvo proizvodi parcelne, agronomske, troškovne i proizvodne podatke;
- Savetnik pretvara te podatke u naloge, preporuke i praćenje izvršenja, i vraća agronomski kvalitet nazad u Gazdinstvo;
- GGAP modul iste podatke pretvara u evidencije, dokaze, zadatke i kontrolu, a njegovi zahtevi podižu disciplinu podataka u Enterprise i Gazdinstvo sistemu.

Ova povezanost je važnija konkurentska prednost od bilo kog pojedinačnog ekrana.

---

## 3. Vizija

Do 2030. AgriX treba da bude vodeća regionalna platforma za:

- kompletno upravljanje organizovanim otkupom;
- digitalno upravljanje gazdinstvima;
- automatizovan GGAP dokumentacioni tok.

AgriX treba da povezuje centralu, stanice, management, administraciju, finansije, Dispečera, vozače, magacin, kupce, agronome, GGAP koordinatore i kooperante u jednom sistemu.

Strateški cilj je najmanje 200 firmi u naredne 3–4 godine. To je ambicija koju treba operacionalizovati kroz tržište, prodajni kapacitet, onboarding, podršku, razvoj i kapital; nije garantovana prognoza.

---

## 4. Misija

AgriX omogućava firmi da vodi ceo poslovni ciklus u jednom povezanom sistemu: od parcele, kooperanta i izdavanja repromaterijala, preko otkupa, dokumentacije, transporta, prijema i lagera, do otpreme, fakture, SEF-a, banke, naplate, isplate i management kontrole.

AgriX Gazdinstvo omogućava proizvođaču da vodi proizvodnju, parcele, tretmane, troškove, opremu, prognozu, bilans i odnos sa hladnjačom.

AgriX GGAP pretvara već nastale operativne podatke u kontrolisan dokumentacioni i dokazni tok bez nepotrebnog ponovnog unosa.

---

## 5. Strateška teza

AgriX može osvojiti značajan deo tržišta samo ako istovremeno ispuni četiri uslova:

1. **Širina:** pokriva ceo posao, a ne samo otkupni list.
2. **Pouzdanost:** kritični tokovi rade u realnim sezonskim uslovima.
3. **Ponovljivost:** onboarding, konfiguracija i podrška ne zavise trajno od osnivača.
4. **Brzina:** readiness se brzo pretvara u prodaju i tržišni udeo.

Previše oprezan rast povećava rizik da konkurent zauzme tržište. Prebrz rast bez readiness-a povećava rizik da AgriX izgubi reputaciju. Strategija zato nije „ostani mali“, već:

> Što brže povećavati readiness i zatim readiness pretvarati u tržišni udeo.

---

## 6. Šta AgriX jeste, a šta nije

### AgriX jeste

- vertikalni poslovni operativni sistem;
- end-to-end sistem za glavne i ključne sporedne tokove otkupnog biznisa;
- farm-management proizvod;
- GGAP documentation i compliance workflow (kao modul Enterprise-a);
- upravljački sloj za agronome i savetodavne službe;
- B2B2C ekosistem koji povezuje firmu i kooperanta;
- zajednički proizvod bez trajnih klijentskih forkova;
- integracioni sloj prema SEF-u, bankama i drugim sistemima;
- potencijalni dobavljač šireg IT sistema ciljnom segmentu.

### AgriX nije

- generički ERP za sve industrije;
- univerzalni knjigovodstveni program;
- custom software studio koji pravi posebnu verziju za svakog klijenta;
- jeftin alat samo za štampanje otkupnih listova;
- read-only portal za kooperante;
- statična arhiva GGAP PDF obrazaca;
- garancija sertifikacije ili zamena za stručnog konsultanta;
- klasičan hardverski distributer bez sopstvene tehnološke vrednosti;
- proizvod koji se prodaje kao production pre prolaska release kriterijuma;
- projekat koji radi rewrite samo zato što je nova tehnologija privlačnija.

---

## 7. Ciljno tržište i model kupca

### 7.1 Geografski fokus

Trenutni fokus je Srbija. Dugoročni cilj je regionalna platforma, uz redosled koji će biti potvrđen tržišnim i regulatornim istraživanjem: BiH, Crna Gora, Severna Makedonija, Hrvatska i druga relevantna tržišta.

### 7.2 Primarni B2B ICP

Najbolji trenutni kupac je firma koja:

- ima razgranatu mrežu otkupnih stanica i kooperanata;
- ima sopstvenu ili organizovanu logistiku;
- izdaje repromaterijal ili ambalažu;
- ima veliki obim dokumenata i finansijskih transakcija;
- želi centralnu management kontrolu;
- koristi postojeći knjigovodstveni program, ali nema dobar operativni sistem;
- ima internog championa i vlasnika implementacije;
- prihvata standardan proizvod i konfiguraciju bez forka;
- ima GGAP ili sličan dokumentacioni pritisak, ili očekuje da će ga imati.

Broj stanica je važan, ali nije dovoljan kriterijum. Širina procesa, broj kooperanata, logistika, repromaterijal, finansije i GGAP mogu učiniti firmu sa manjim brojem stanica vrednijim klijentom.

### 7.3 Buying committee

Prodaja mora obuhvatiti najmanje:

- ekonomskog kupca: vlasnik ili direktor;
- championa: administrator, manager ili mlađi vlasnik;
- ključnog operativnog korisnika: otkupljivač, Dispečer, finansije ili magacioner;
- mogućeg blockera;
- tehničkog i compliance influensera kada su relevantni.

Ponuda ne treba da se zasniva samo na razgovoru sa vlasnikom. Pre pune implementacije mora biti potvrđena operativna realnost najmanje tri ključne uloge.

### 7.4 Gazdinstvo ICP

Najbolji korisnik ima više parcela ili intenzivnu proizvodnju, koristi agrohemiju, želi kontrolu troškova i rezultata, sarađuje sa AgriX hladnjačom i ima motivaciju da vodi evidenciju.

### 7.5 GGAP ICP

Najbolji kupac je firma ili grupa proizvođača koja upravlja dokumentacijom većeg broja gazdinstava, ima imenovanog quality/GGAP ownera i želi kontinuiranu spremnost, a ne jednokratno sređivanje fascikli pred kontrolu.

### 7.6 Anti-ICP

AgriX ne treba aktivno da prihvata klijenta koji:

- zahteva trajni fork ili poseban release;
- očekuje neograničen custom development uključen u licencu;
- nema odgovornu osobu za podatke i implementaciju;
- traži rollout neposredno pred sezonu bez pilota i testa;
- odbija backup, monitoring, update ili bezbednosna pravila;
- odbija obuku korisnika;
- kupuje isključivo po najnižoj ceni;
- očekuje garanciju poslovnog rezultata ili GGAP sertifikacije;
- ne prihvata podelu odgovornosti za hardver, internet i kvalitet podataka;
- pokazuje visok rizik neplaćanja ili zloupotrebe supporta.

Detaljan ICP scoring i jobs-to-be-done nalaze se u `03_CUSTOMERS_AND_JOBS.md`.

---

## 8. Strateški principi proizvoda

### 8.1 Jedan kod, bez forkova

Razlike među firmama rešavaju se konfiguracijom, modulima, rolama, dozvolama, workflow pravilima i feature flags. Trajni klijentski fork nije dozvoljen.

### 8.2 Pokriti posao, ne samo funkcije

Funkcionalnost ima stratešku vrednost kada zatvara poslovni tok, uklanja ručni prelaz ili povećava kontrolu. Duga lista nepovezanih funkcija nije dovoljna.

### 8.3 Jedinstven podatak kroz ceo tok

Podatak nastao na parceli, u magacinu, na stanici ili u banci koristi se dalje bez ponovnog ručnog unosa. Dupli unos je signal da tok nije potpuno zatvoren.

### 8.4 Dokument nastaje iz operacije

GGAP, finansijski i operativni dokumenti treba da koriste izvorne događaje i da zadrže provenance: ko, kada, gde, iz kog zapisa i uz koje kasnije korekcije.

### 8.5 Mobilni uređaj je operativni terminal

Kamera, QR, barkod, digitalni potpis, GIS, fotografija, geolokacija i PWA služe da transakcija i dokaz nastanu na mestu događaja.

### 8.6 Podatak po parceli je strateška imovina

Povezivanje parcele sa prognozom, tretmanima, troškovima, proizvodnjom, otkupom, rezultatom i GGAP dokazima stvara odbranjivu vrednost.

### 8.7 Compliance sadržaj mora biti verzionisan

GGAP liste, pravila i validacije moraju imati verziju, period važenja i jasno vlasništvo. Nova verzija ne sme nevidljivo menjati istorijske dokumente.

### 8.8 Softver podržava, ali ne garantuje usklađenost

AgriX vodi workflow, proverava podatke, upozorava i priprema dokaze. Sertifikacija i stručna odluka ostaju van domena softverske garancije.

### 8.9 Operativna jednostavnost je funkcionalnost

Onboarding, monitoring, self-update, backup, recovery, kiosk konfiguracija, manuali i runbook-ovi imaju isti strateški značaj kao korisničke funkcije.

### 8.10 Bez rewrite-a bez merljivog razloga

Promena platforme razmatra se tek kada postojeća tehnologija stvara merljiv limit u pouzdanosti, brzini razvoja, zapošljavanju, integracijama ili trošku održavanja.

### 8.11 Staged rollout

Velike promene prolaze kroz interni test, pilot, ograničenu grupu i tek zatim pun rollout.

### 8.12 Ne prodavati planirano kao postojeće

Svaki domen mora imati status `Production`, `Pilot`, `Planned`, `Gap` ili `Out of scope`. Komercijalna tvrdnja mora odgovarati stvarnom statusu.

---

## 9. Fiksan ciljni broj klijenata

`DECISION` (odluka 403): rast se planira prema **fiksnom ciljnom broju klijenata**, ne prema readiness cap-u. Ovim se povlači STR-001.

Aktuelan cilj je scenario C iz odluke 375: **14 novih Enterprise klijenata do sezone 2027, ukupno 17 aktivnih firmi**. To je jedna planska vrednost, ne raspon. Odluka 375 zamenjuje odluku 249. Isti broj stoji u `04_MARKET.md` §9.1 i u `docs/Finance/AgriX_Finansijski_model.xlsx` (list `Pretpostavke`, red 42); ako se cilj promeni, menjaju se sva tri mesta.

Kolona od 18 klijenata na listu `Kapacitet` je stress-test koji pokazuje kada osnivač postaje usko grlo — nije cilj i ne koristi se u planiranju prihoda.

Šta se ovom odlukom menja, a šta ne:

- **menja se:** ciljni broj klijenata se određuje unapred i ne pomera ga readiness score;
- **ne menja se:** readiness ostaje obavezan preduslov kvaliteta isporuke. Crven P0 domen i dalje može zaustaviti ili odložiti **pojedinačan** onboarding, ali više ne obara ciljni broj — odgovor na crven domen je otklanjanje uzroka ili dodavanje kapaciteta, ne spuštanje cilja.

Zbog toga readiness lista prelazi iz uloge cap mehanizma u operativnu kontrolnu listu pred svaki onboarding. Mora najmanje obuhvatiti:

- stabilnost kritičnog codebase-a;
- production readiness ugovorenih poslovnih tokova;
- Field, Vozač, Dispečer, repromaterijal i Gazdinstvo;
- QR/barkod, GIS, meteo, offline sync i štampu;
- SEF, banku i finansijski reconciliation;
- GGAP sadržaj, validacije, dokazni tok i export kada je GGAP modul aktiviran;
- trajanje i automatizaciju onboardinga;
- kvalitet manuala, checklista i runbook-ova;
- support kapacitet i eskalacije;
- monitoring, backup, recovery i release proces;
- logistiku hardvera i rezervnu opremu;
- finansijsku rezervu i obrtni kapital;
- broj osoba koje mogu sprovesti onboarding bez osnivača.

Visok prosečan score ne može sakriti kritičnu slabost. Jedan crveni P0 domen blokira onboarding klijenata koje taj domen dodiruje, bez obzira na ostale rezultate — ali je odgovor otklanjanje uzroka, a ne smanjenje ciljnog broja iz odluke 403.

---

## 10. Strategija rasta 2026–2030

### Faza 1 — Dokaz ponovljivosti i readiness modela

Ciljevi:

- standardizovati remote onboarding;
- pripremiti manuale i checkliste koje koristi support/implementation osoba;
- potvrditi kritične Enterprise tokove u realnoj sezoni;
- definisati production status svih Gazdinstvo funkcija;
- sprovesti GGAP discovery pre pune implementacije;
- meriti onboarding, support i incidente po firmi i domenu;
- potvrditi da rast ne zahteva forkove i ne povećava incident rate po firmi;
- dovesti readiness na nivo koji podržava fiksni cilj, umesto da cilj prilagođava readiness-u.

Ciljni broj firmi je fiksiran odlukom 403 — scenario C iz odluke 375, odnosno 14 novih klijenata i ukupno 17 aktivnih Enterprise firmi do sezone 2027. Readiness više ne određuje taj broj; on određuje da li je konkretan onboarding bezbedan i koliko kapaciteta treba dodati da bi cilj bio dostižan.

### Faza 2 — Ubrzana nacionalna penetracija

**Okvir:** približno 10–50 firmi.

- customer support / implementation preuzima standardan onboarding i bazne slučajeve;
- osnivač ostaje eskalacija, product owner i ključni prodavac;
- osnivač se postepeno prebacuje sa operacije na tržište;
- developer se dodaje kada razvoj postane dokazano usko grlo;
- grade se case studies, reference, direktna prodaja i partnerstva;
- potvrđuju se pricing i unit economics Enterprise i Gazdinstvo proizvoda;
- GGAP prolazi kroz ograničeni pilot.

### Faza 3 — Liderstvo u Srbiji

**Okvir:** približno 50–200 firmi.

- izgraditi najprepoznatljiviji specijalizovani brend za kompletan otkupni biznis;
- organizovati support, implementaciju i razvoj bez dnevne zavisnosti od osnivača;
- standardizovati hardverski i širi IT katalog;
- razviti partnerstva sa bankama, knjigovođama, agronomima, GGAP konsultantima i dobavljačima opreme;
- potvrditi ili odbaciti Gazdinstvo i GGAP ekonomiku;
- pripremiti lokalizaciju i operativni model za region.

### Faza 4 — Regionalna platforma

Regionalno širenje zahteva:

- tržišno i regulatorno mapiranje po zemlji;
- lokalizaciju jezika, dokumenata, poreza, banke i e-faktura;
- lokalni prodajni i implementation kanal;
- definisan support model;
- jasno vlasništvo nad lokalnim GGAP/compliance sadržajem.

---

## 11. Strategija prihoda

Potencijalni izvori prihoda:

1. godišnje licence za AgriX Enterprise;
2. paketi prema širini procesa i operativnom obimu;
3. implementacija, migracija i obuka;
4. multi-company licence, premium support i SLA;
5. hardver i širi IT sistem sa stvarnom pozitivnom maržom;
6. Gazdinstvo Partner, Basic i Pro;
7. AgriX GGAP licence i dokumentacioni paketi;
8. buduće integracije i premium usluge.

Pricing ne treba da pretvori AgriX u konfuznu listu desetina mikro-doplata. Packaging mora da sačuva vrednost povezanog sistema, uz jasne pakete prema broju stanica, kooperanata, korisnika i širini procesa.

### 11.1 Enterprise

Enterprise trenutno finansira osnovni biznis. Cena mora odražavati poslovnu kritičnost, širinu toka, broj stanica, implementaciju, support i pun trošak održavanja — ne samo broj desktop korisnika.

### 11.2 Gazdinstvo

Konzervativni finansijski model ne sme pretpostaviti značajan prihod dok aktivacija i konverzija nisu potvrđene. Dugoročni potencijal ostaje otvoren.

### 11.3 GGAP

Mogući modeli su licenca po firmi ili grupi, osnovna cena plus aktivna gazdinstva, samostalni paket za proizvođača i odvojena migracija/onboarding usluga. Trošak stručnog održavanja sadržaja mora ući u unit economics.

GGAP se ne ceni kao generator PDF-a, već prema vrednosti kontinuirane spremnosti, smanjenju ručnog rada i ranom otkrivanju propusta.

### 11.4 Hardver i širi IT sistem

Hardver nije glavni profitni centar, ali mora biti profitabilan sporedni centar. AgriX može postati dobavljač kiosk tableta, termalnih štampača, uređaja sa pouzdanom kamerom, računara, mrežne i rezervne opreme, remote managementa i integracije periferija.

Svaka kategorija mora imati prikaz prodajne cene, nabavne cene, rada, transporta, garancije, zamene, zalihe i obrtnog kapitala.

---

## 12. Strategija organizacije

### 12.1 Osnivač

U narednoj fazi osnivač zadržava:

- product ownership;
- arhitekturu i ključni razvoj;
- finalnu eskalaciju;
- prodaju važnim klijentima;
- marketing strategiju;
- ključna partnerstva.

Cilj je da osnivač prestane da bude usko grlo za standardni onboarding i support.

### 12.2 Prvo zaposlenje

Prva operativna osoba je customer support / implementation.

Odgovornosti:

- rešavanje baznih pitanja;
- onboarding prema manualima i checklistama;
- pomoć oko konfiguracije, tableta, kamera, barkoda i štampača;
- praćenje monitoringa;
- trijaža po domenu i roli;
- rešavanje poznatih slučajeva prema runbook-u;
- eskalacija složenih problema;
- evidencija vremena, uzroka i rešenja.

### 12.3 Developer

Developer se dodaje kada razvoj postane dokazano usko grlo, roadmap kasni ili osnivač mora značajno da se prebaci na prodaju i tržište.

### 12.4 GGAP domain owner

Pre production lansiranja GGAP-a mora biti jasno ko poseduje:

- mapiranje standarda i verzija;
- sadržaj listi i procedura;
- validaciona pravila;
- stručnu reviziju promena;
- odobravanje compliance release-a;
- pitanja klijenata koja nisu tehnički support.

Ova odgovornost ne sme neformalno pasti na customer support osobu.

---

## 13. Partner i kapital

AgriX ne treba partnera samo zbog novca.

Partner ima smisla kada donosi najmanje jednu teško zamenljivu sposobnost:

- direktan pristup velikom broju kvalitetnih kupaca;
- dokazanu distribuciju u agraru;
- vođenje prodaje i implementacije;
- iskustvo skaliranja B2B/B2B2C softvera;
- regionalnu mrežu;
- stručnu GGAP/compliance sposobnost;
- kapital vezan za precizan plan uklanjanja dokazanog uskog grla.

Pre prodaje udela moraju biti poznati upotreba kapitala, očekivani dodatni ARR, rok, odgovornost partnera, upravljačka prava, dilution i scenario neuspeha.

---

## 14. Strateška hitnost i moat

`HYPOTHESIS`: tržište ima ograničen prozor u kojem AgriX može izgraditi dominantnu poziciju pre nego što generički ERP dobavljač ili novi vertikalni konkurent razvije dovoljno sličan sistem.

AgriX-ov moat ne treba zasnivati samo na broju funkcija. Najjače odbrane su:

1. end-to-end tok kroz ceo posao;
2. jedan povezani podatak kroz Enterprise, Gazdinstvo i GGAP;
3. parcelni, agronomski, finansijski i dokumentacioni kontekst;
4. production iskustvo i poslovna pravila specifična za otkup;
5. monitoring, self-update, onboarding i operativni runbook-ovi;
6. mreža firmi, kooperanata i uređaja;
7. switching cost koji nastaje kada AgriX postane primarni operativni sistem.

Širina bez pouzdanosti nije moat. Pouzdanost bez tržišne penetracije takođe nije dovoljna.

---

## 15. Strateški rizici

| Rizik | Verovatnoća | Uticaj | Primarna zaštita |
|---|---|---|---|
| Osnivač ostaje jedina osoba koja razume ceo sistem | Visoka | Kritičan | dokumentacija, support, developer, ownership |
| Širina tri proizvoda prevaziđe kapacitet malog tima | Visoka | Kritičan | jasne faze, statusi, scope i release gate |
| Funkcije postoje, ali nisu spojene u zatvoren tok | Srednja | Visok | end-to-end process mapping |
| Dispečer, Vozač, Field ili štampa nisu dovoljno stabilni | Srednja | Kritičan | pilot, staged rollout i fallback |
| SEF ili banka proizvedu finansijski pogrešan rezultat | Srednja | Kritičan | validation, fail-closed, audit i reconciliation |
| Gazdinstvo ima veliku širinu, ali nisku aktivaciju | Visoka | Visok | analytics, onboarding i test monetizacije |
| GIS/meteo preporuka bude shvaćena kao stručna garancija | Srednja | Visok | izvori, ograničenja i upozorenja |
| GGAP sadržaj ne prati važeću verziju standarda | Srednja | Kritičan | domain owner, verzionisanje i stručna revizija |
| Klijent shvati softver kao garanciju sertifikacije | Srednja | Kritičan | UX, ugovor, edukacija i audit trag |
| Automatski dokaz koristi pogrešan izvorni podatak | Srednja | Kritičan | provenance, approval i validacija |
| Readiness preceni kapacitet | Srednja | Kritičan | weakest-link kontrolna lista pred onboarding i rezerva |
| Cena bude niža od punog troška sistema | Srednja | Visok | unit economics i value-based pricing |
| Hardver veže previše kapitala | Srednja | Visok | predujam, standardni modeli i ograničena zaliha |
| Rast bude prespor i konkurent zauzme tržište | Srednja/visoka | Kritičan | ambiciozan GTM i rast readiness-a |
| Agresivan rast pogorša kvalitet | Srednja | Kritičan | fiksni cilj uz readiness kontrolnu listu i staged onboarding |

---

## 16. Ključni strateški KPI-jevi

### Enterprise

- aktivne firme i nove firme po sezoni;
- readiness score ukupno i po domenu;
- onboarding sati i procenat onboardinga bez osnivača;
- support sati i incidenti po firmi i domenu;
- procenat potpuno zatvorenih poslovnih tokova;
- broj ručnih prelaza i duplih unosa;
- aktivni Field i Driver terminali;
- Dispečer korišćenje i uspešnost planova;
- uspešnost sync-a i štampe;
- SEF uspešnost i neusaglašeni statusi;
- procenat automatski rasknjiženih bankarskih stavki;
- pripremljeni nalozi za plaćanje;
- ARR, gross margin, renewal i churn;
- stvarna hardverska marža.

### Gazdinstvo

- Partner, Basic i Pro nalozi;
- aktivacija po hladnjači;
- WAU/MAU i retencija;
- aktivne parcele;
- GIS/meteo korišćenje;
- evidentirani tretmani i korišćenje pametnog doziranja;
- uneti troškovi i procenat vezan za parcelu;
- pregledi bilansa i kartice prema hladnjači;
- Partner → Basic/Pro konverzija;
- ARPU i support cost.

### GGAP

- aktivne firme, grupe i gazdinstva;
- procenat automatski popunjenih polja i listi;
- procenat kompletiranosti dokumentacije;
- nedostajući i istekli dokazi;
- vreme pripreme dokumentacije po gazdinstvu;
- otvorene i zatvorene neusaglašenosti;
- vreme zatvaranja korektivne mere;
- broj audit paketa;
- propusti pronađeni pre kontrole;
- ARR, gross margin, renewal i support/content cost.

---

## 17. Odobrene strateške odluke

- **STR-002:** trenutni fokus je Srbija i firme sa razgranatom mrežom stanica i kooperanata;
- **STR-003:** jedan proizvodni codebase, bez trajnih klijentskih forkova;
- **STR-004:** Gazdinstvo trenutno nije osnovni prihod, ali može postati glavni proizvod ili prihod;
- **STR-005:** hardver je profitabilan sporedni centar i mogući ulaz u širi IT portfolio;
- **STR-006:** partner se ne uzima samo zbog kapitala;
- **STR-007:** prvo operativno zaposlenje je customer support / implementation;
- **STR-008:** dugoročni cilj je regionalna platforma;
- **STR-009:** cilj je najmanje 200 firmi u naredne 3–4 godine;
- **STR-010:** AgriX pokriva ceo poslovni sistem firme;
- **STR-011:** Gazdinstvo je pun farm-management proizvod;
- **STR-013:** treći proizvodni stub je AgriX Savetnik (odluka 401);
- **STR-014:** rast se planira prema fiksnom ciljnom broju klijenata (odluka 403).

### Povučene odluke

| Odluka | Status | Zamenjena sa |
|---|---|---|
| STR-001 — readiness-based rast | `Superseded` 2026-07-27 | odluka 403 → STR-014 |
| STR-012 — GGAP kao treći proizvodni stub | `Superseded` 2026-07-27 | odluke 401 i 402 → STR-013; GGAP je modul Enterprise-a |

Detaljni razlozi i posledice odluka vode se u `DECISION_LOG.md`, a numerisane odluke u `09_QA_DECISION_LOG.md`.

---

## 18. Otvorene teme

1. potvrditi veličinu tržišta i segmentaciju u `04_MARKET.md`;
2. izraditi `07_PRODUCT_PORTFOLIO.md` sa statusom svakog toka;
3. definisati readiness kontrolnu listu pred onboarding (više nije cap mehanizam — odluka 403);
4. razložiti cilj od 200 firmi na godišnji prodajni, kadrovski i finansijski plan;
5. završiti GGAP discovery kao preduslov za izlazak modula iz pilota (odluka 405);
6. definisati Gazdinstvo Partner, Basic i Pro granice;
7. definisati packaging Enterprise-a, Gazdinstva i Savetnika, uključujući module Enterprise-a;
8. fino podesiti unit economics po proizvodu, paketu i segmentu — posle određenih cena, ne kao preduslov (odluka 408);
9. odrediti IT kategorije koje AgriX prodaje i podržava;
10. definisati pragove za partnera ili investitora.

---

## 19. Naredni koraci

1. razviti `04_MARKET.md` i potvrditi adresabilno tržište;
2. razviti `07_PRODUCT_PORTFOLIO.md` kao jedinstvenu mapu tri stuba, modula Enterprise-a i svih poslovnih tokova;
3. napraviti readiness kontrolnu listu pred onboarding;
4. sprovesti intervjue i validaciju iz `03_CUSTOMERS_AND_JOBS.md`;
5. razviti Savetnik product strategy kao trećeg stuba (odluka 401);
6. uneti odluke 322–400 u `09_QA_DECISION_LOG.md`, jer se na njih poziva više dokumenata;
7. zatim razviti pricing dokumentaciju, finansijski model i plan rasta do 200 firmi.
