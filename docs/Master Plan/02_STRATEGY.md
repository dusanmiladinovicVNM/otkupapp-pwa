# 02 — Strategija i identitet AgriX-a

**Status:** Review  
**Vlasnik:** osnivač AgriX-a  
**Horizont:** 2026–2030  
**Poslednje ažuriranje:** 2026-07-22

---

## 1. Polazna tačka

AgriX je specijalizovan poslovni sistem za organizaciju otkupa poljoprivrednih proizvoda. Trenutni proizvod povezuje desktop administraciju, management PWA, monitoring, self-update i niz poslovnih modula. Za narednu sezonu planirani su usklađeni PWA Otkupac, kiosk tableti i termalna štampa otkupnih listova na mestu otkupa.

Potvrđene polazne činjenice:

- `FACT`: postoje tri aktivna klijenta;
- `TARGET`: približno pet novih klijenata naredne sezone, ukupno oko osam;
- `TARGET`: preko deset firmi nije cilj naredne sezone;
- `FACT`: prosečna firma ima približno deset otkupnih stanica;
- `FACT`: prosečno postoji jedan desktop korisnik po firmi, uz management PWA;
- `FACT`: prosečno postoji oko 100 kooperanata po firmi;
- `MEASURED`: support je približno jedan poziv nedeljno u dosadašnjem periodu;
- `FACT`: self-update i granularni monitoring postoje;
- `FACT`: onboarding je do sada rađen remote;
- `FACT`: sve klijentske varijacije ostaju u zajedničkom kodu i rešavaju se konfiguracijom.

Ove činjenice pokazuju da AgriX više nije prototip, ali još nije dokazano skaliran proizvod. Sledeća sezona mora potvrditi da novi terenski sloj radi pouzdano na većem broju stanica i firmi.

---

## 2. Strateška definicija

AgriX nije generički ERP i nije zbir nepovezanih modula. AgriX je vertikalna operativna platforma za firme koje organizuju otkup i za ljude koji učestvuju u tom toku.

Platforma ima tri osnovna sloja:

1. **AgriX Otkup / Desktop** — centralna administracija, dokumentacija, finansijski i operativni tokovi;
2. **AgriX Field / PWA Otkupac** — rad na otkupnim stanicama, kiosk terminali i štampa na licu mesta;
3. **AgriX Gazdinstvo** — digitalna veza kooperanta sa hladnjačom i samostalna evidencija gazdinstva.

Hladnjača je primarni kupac i glavni izvor prihoda u kratkom i srednjem roku. Gazdinstvo je u početku dodatak ekosistemu i kanal za vezivanje kooperanata, a ne glavni finansijski oslonac.

---

## 3. Vizija

Do 2030. AgriX treba da bude najrelevantnija specijalizovana platforma za digitalizaciju otkupa kod malih i srednjih otkupljivača i hladnjača u Srbiji, sa dokazanim modelom koji se može preneti na odabrana tržišta regiona.

Vizija ne znači najveći broj funkcija niti najveći broj instalacija. Znači da AgriX postane proizvod koji:

- pouzdano vodi kritični tok otkupa tokom sezone;
- povezuje centralnu administraciju sa terenskim stanicama;
- smanjuje ručni rad, prepisivanje i kašnjenje informacija;
- daje managementu trenutni uvid i kontrolu;
- ostaje standardizovan proizvod sa jednim kodom;
- može da se održava sa malim, kvalitetnim timom;
- stvara dovoljno recurring prihoda da razvoj i podrška ne zavise isključivo od osnivača.

---

## 4. Misija

AgriX omogućava domaćim otkupljivačima da vode otkup, dokumentaciju, ambalažu, finansijske tokove i terenske stanice kao jedan povezan sistem, bez troška i složenosti velikog ERP projekta.

Za kooperante AgriX treba da obezbedi jasan digitalni pregled saradnje sa hladnjačom i jednostavan alat za sopstvenu evidenciju gazdinstva.

---

## 5. Šta AgriX namerno nije

AgriX namerno nije:

- univerzalni knjigovodstveni program;
- potpuna zamena za BizniSoft, PANTHEON ili drugi računovodstveni ERP;
- custom software studio koji pravi poseban proizvod za svakog klijenta;
- jeftin program samo za štampanje otkupnih listova;
- hardverski distributer kome je prodaja tableta osnovni biznis;
- platforma koja agresivno prihvata klijente pre nego što može bezbedno da ih podrži;
- projekat koji menja tehnologiju samo zato što je nova tehnologija atraktivnija;
- proizvod koji obećava funkcionalnosti koje nisu production-ready.

Ova ograničenja su strateška zaštita fokusa.

---

## 6. Idealni kupac u prvoj fazi

Primarni idealni kupac je mala ili srednja hladnjača ili otkupljivač u Srbiji sa sledećim karakteristikama:

- godišnji promet približno 1–2 miliona EUR;
- više otkupnih stanica, tipično oko deset;
- jedan centralni administrativni korisnik;
- vlasnik ili management koji želi pregled i kontrolu preko PWA;
- postojeći računovodstveni sistem koji ne rešava dobro ulazni tok robe;
- spremnost da standardizuje proces umesto da zahteva poseban fork;
- dovoljno veliki operativni problem da godišnja licenca ima jasnu vrednost;
- dovoljno mala organizacija da odluku može doneti vlasnik ili direktor bez višemesečne nabavke.

---

## 7. Anti-ICP

AgriX ne treba aktivno da prihvata klijente koji:

- zahtevaju sopstvenu verziju koda ili poseban release;
- očekuju neograničen custom development uključen u licencu;
- nemaju odgovornu osobu za podatke, obuku i komunikaciju;
- traže produkcijski rollout neposredno pred sezonu bez vremena za test;
- odbijaju standardne procese backupa, update-a i monitoringa;
- kupuju isključivo po najnižoj ceni;
- imaju toliko kompleksan enterprise proces da zahtevaju poseban implementacioni tim i SLA koji AgriX još nema;
- očekuju da AgriX preuzme fizičke kvarove, lom i zloupotrebu hardvera bez ugovorne granice.

Odbijanje lošeg klijenta može biti profitabilnija odluka od prodaje.

---

## 8. Strateški principi razvoja

### 8.1 Jedan kod, mnogo konfiguracije

Sve funkcionalnosti ostaju u zajedničkom kodu. Razlike među firmama rešavaju se kroz podešavanja, module, dozvole, workflow konfiguraciju i feature flags.

Zabranjeni obrazac je trajna logika vezana za ime ili identitet pojedinačnog klijenta.

### 8.2 Pouzdanost pre širine

Funkcionalnost koja vodi kritičnu transakciju mora biti stabilna, merljiva i recoverable pre nego što se doda nova širina proizvoda.

### 8.3 Kontrolisan rollout

Velike promene prolaze kroz interni test, pilot firmu, ograničenu grupu i tek zatim pun rollout. Self-update omogućava distribuciju, ali ne uklanja potrebu za staged rolloutom.

### 8.4 Bez rewrite-a bez merljivog razloga

Promena tehnološke platforme razmatra se kada postojeća platforma stvara merljiv limit u pouzdanosti, brzini razvoja, zapošljavanju, integracijama ili ukupnom trošku održavanja. Rewrite nije strateški cilj sam po sebi.

### 8.5 Operativna jednostavnost je funkcionalnost

Remote onboarding, monitoring, self-update, backup, kiosk konfiguracija i jasni runbook-ovi imaju isti strateški značaj kao korisničke funkcije.

### 8.6 Nema prikrivenog custom developmenta

Zahtev jednog klijenta ulazi u proizvod samo kada predstavlja opšti problem segmenta i može se rešiti kroz zajednički model. Poseban razvoj se posebno ugovara ili odbija.

### 8.7 Ne obećavati budući proizvod kao postojeći

PWA Otkupac, termalna štampa, kiosk terminali i budući moduli prodaju se kao production tek kada zadovolje release kriterijume.

---

## 9. Strategija rasta 2026–2030

### Faza 1 — Dokaz operativne skale

**Period:** naredna sezona  
**Cilj:** ukupno približno 8 firmi, maksimalno 10.

Primarni cilj nije maksimalan prihod, već dokaz da:

- novi klijent može biti onboardovan remote za pola dana;
- PWA Otkupac ostaje usklađen sa desktop tokom;
- kiosk terminali i termalna štampa rade pouzdano;
- support ostaje merljiv i podnošljiv;
- update i monitoring rade na svim firmama;
- nema potrebe za klijentskim fork-ovima;
- pricing može da se proda bez velikog popusta.

### Faza 2 — Standardizacija i delegiranje

**Okvir:** približno 10–20 firmi.

Ciljevi:

- customer support / implementation osoba preuzima većinu standardnih pitanja i onboardinga;
- osnivač ostaje eskalacija za bugove i poslovnu logiku;
- svi standardni procesi imaju checklistu i runbook;
- stvarni support cost i onboarding cost ulaze u unit economics;
- dodatni developer se angažuje samo kada je razvoj dokazano usko grlo;
- marketing postaje sistematska funkcija, a ne povremena aktivnost.

### Faza 3 — Održiv rast

**Okvir:** približno 20–50 firmi.

Ciljevi:

- recurring prihod finansira minimalni stalni tim;
- podrška i implementacija ne zavise svakodnevno od osnivača;
- postoji pouzdan release i incident proces;
- Gazdinstvo ima dokazanu, ne pretpostavljenu konverziju;
- regionalno širenje počinje tek kada je domaći model standardizovan;
- cena raste sa dokazima, referencama i višim nivoom usluge.

### Faza 4 — Platforma ili profitabilni specijalista

Do 2030. AgriX može izabrati jedan od dva zdrava ishoda:

1. profitabilna specijalizovana firma sa ograničenim timom i visokim kvalitetom usluge;
2. regionalna vertikalna platforma sa većim timom, dodatnim kapitalom i širim proizvodnim portfoliom.

Odluka se ne donosi ideološki. Donosi se na osnovu ARR-a, tržišne tražnje, kapaciteta tima, churn-a i stvarne ekonomike Gazdinstva.

---

## 10. Strategija prihoda

Redosled važnosti izvora prihoda u narednim godinama:

1. godišnje licence hladnjača;
2. implementacija i obuka kao odvojena usluga;
3. dodatni moduli i multi-company licence;
4. realna marža na konfiguraciji i isporuci terminala;
5. Gazdinstvo Basic i Pro;
6. buduće integracije i premium SLA.

Hardver nije centralni profitni motor. Njegova uloga je da omogući pouzdan terenski rad i poveća vrednost softverske platforme.

Gazdinstvo je u prvoj fazi strateški proizvod, ali finansijski trivijalan. Sa približno 100 kooperanata po firmi i radnom hipotezom konverzije od 5%, njegov prihod ne sme da finansira osnovni tim dok tržište to ne potvrdi.

---

## 11. Strategija organizacije

### Osnivač

U narednoj fazi osnivač zadržava:

- product ownership;
- arhitekturu i ključni razvoj;
- finalnu eskalaciju supporta;
- prodaju važnim klijentima;
- marketing strategiju;
- ključna partnerstva.

Ovo je privremeno održivo do planirane granice sledeće sezone, ali nije ciljna organizacija za 20–50 firmi.

### Prvo zaposlenje

Planirana prva operativna osoba je customer support / implementation. Njena uloga uključuje:

- standardna korisnička pitanja;
- remote onboarding;
- pomoć oko tableta, štampača i konfiguracije;
- praćenje monitoringa;
- evidenciju i trijažu problema;
- eskalaciju bugova osnivaču.

### Dodatni developer

Developer se dodaje kada razvoj i održavanje postanu dokazano usko grlo ili kada osnivač prelazi značajnije na marketing i prodaju. Developer ne sme biti zaposlen samo zato što rast izgleda poželjno.

---

## 12. Partner i kapital

AgriX trenutno ne treba partnera samo zbog kapitala. Kapital bez jasnog uskog grla može povećati troškove i pritisak bez proporcionalnog rasta.

Partner ima smisla kada donosi najmanje jednu teško zamenljivu sposobnost:

- direktan pristup velikom broju kvalitetnih kupaca;
- dokazanu distribuciju u agraru;
- operativno vođenje prodaje i implementacije;
- relevantno iskustvo skaliranja B2B softvera;
- regionalnu mrežu koju AgriX ne može brzo sam da izgradi;
- kapital vezan za precizan plan koji je već validiran.

Pre prodaje udela moraju biti poznati:

- tačna upotreba kapitala;
- očekivani dodatni ARR;
- rok do rezultata;
- odgovornost partnera;
- upravljačka prava;
- dilution i izlazni scenariji;
- šta se dešava ako plan ne uspe.

`HYPOTHESIS`: racionalniji trenutak za ozbiljno razmatranje partnera je nakon pune sezone sa 8–10 firmi i dokazanim Field sistemom.

---

## 13. Geografska strategija

Redosled:

1. Srbija;
2. BiH;
3. Crna Gora i Severna Makedonija;
4. Hrvatska ili druga tržišta tek nakon pravne, jezičke i računovodstvene procene.

Regionalno širenje ne počinje samo zato što je proizvod tehnički dostupan. Potrebni su lokalni partner, prodajni kanal, regulatorno mapiranje i jasan support model.

---

## 14. Strateški rizici

| Rizik | Verovatnoća | Uticaj | Primarna zaštita |
|---|---|---|---|
| Osnivač ostaje jedina osoba koja razume ceo sistem | Visoka | Visok | dokumentacija, support osoba, code ownership |
| Previše razvoja pred sezonu | Visoka | Visok | scope freeze i release gate |
| Field štampa ili sync nisu dovoljno pouzdani | Srednja | Kritičan | pilot, staged rollout, rezervni proces |
| Cena je niža od punog troška usluge | Srednja | Visok | unit economics po firmi |
| Hardver veže previše kapitala | Srednja | Srednji/visok | predujam, standardni modeli, mala zaliha |
| Klijenti traže prikriven custom razvoj | Visoka | Srednji | ugovorne granice i product pravila |
| Gazdinstvo nema dovoljnu konverziju | Visoka | Srednji | tretirati prihod kao nulu u konzervativnom planu |
| Agresivan rast pogorša kvalitet | Srednja | Kritičan | limit broja novih firmi |
| Tehnološki dug uspori razvoj | Srednja | Visok | merljivi pragovi i kontrolisan refactoring |
| Partnerstvo prerano smanji kontrolu i buduću vrednost | Srednja | Visok | jasni kriterijumi pre prodaje udela |

---

## 15. Ključni strateški KPI-jevi

Za narednu sezonu:

- broj aktivnih hladnjača: cilj 8, hard cap 10;
- onboarding: cilj najviše 4 sata standardnog rada po firmi;
- kritični incidenti: cilj 0 nerecoverable događaja;
- prosečno support vreme po firmi mesečno;
- procenat problema rešenih bez osnivača;
- broj uspešnih self-update rollouta;
- broj aktivnih Field terminala;
- stopa neuspešne ili ponovljene štampe;
- broj klijentskih zahteva rešenih konfiguracijom naspram novog koda;
- renewal namera i ostvareni renewal;
- ostvarena cena naspram cenovnika;
- onboarding i support cost po firmi;
- stvarna konverzija Gazdinstva, odvojeno od partner naloga.

---

## 16. Preporučene odluke za odobravanje

### STR-001 — Kontrolisan rast

Naredna sezona je sezona dokazivanja operativne skale. Cilj je približno osam firmi, uz maksimalno deset.

### STR-002 — Primarni tržišni fokus

Primarni fokus ostaje Srbija i firme sa više otkupnih stanica kojima generički ERP ne rešava ulazni tok robe.

### STR-003 — Jedan proizvod, jedan kod

Klijentske razlike rešavaju se zajedničkim kodom i konfiguracijom. Trajni klijentski fork nije dozvoljen.

### STR-004 — Hladnjače finansiraju osnovni biznis

Godišnje licence hladnjača moraju finansirati osnovnu organizaciju. Gazdinstvo se u kratkoročnom planu ne tretira kao ključni prihod.

### STR-005 — Hardver je enablement, ne centralni profitni centar

Terminali se nude da bi Field proizvod radio pouzdano. Finansijski se meri stvarna marža nakon svih troškova i rizika.

### STR-006 — Bez partnera samo zbog novca

Partner ili investitor se razmatra tek kada rešava dokazano usko grlo i donosi merljivu sposobnost pored kapitala.

### STR-007 — Prvo operativno zaposlenje

Prva planirana operativna uloga je customer support / implementation. Dodatni developer se angažuje prema razvojnom uskom grlu i prelasku osnivača ka marketingu.

---

## 17. Otvorena pitanja za review

1. Da li je formulacija vizije do 2030. dovoljno ambiciozna ili preširoka?
2. Da li hard cap od deset firmi treba da bude apsolutan ili uslovljen readiness score-om?
3. Da li customer support / implementation osoba treba da bude angažovana pre početka sezone ili tek nakon prelaska određenog broja firmi?
4. Da li AgriX želi da dugoročno ostane profitabilni specijalista ili aktivno cilja regionalnu platformu?
5. Koji konkretan ARR i operativni KPI moraju biti ispunjeni pre razgovora sa partnerom ili investitorom?

---

## 18. Naredni koraci

1. pregledati i odobriti ili izmeniti STR-001 do STR-007;
2. odobrene odluke upisati u `DECISION_LOG.md`;
3. zaključati osnovnu viziju i granicu rasta;
4. razviti `03_CUSTOMERS_AND_JOBS.md`;
5. razviti `07_PRODUCT_PORTFOLIO.md`;
6. tek zatim finalizovati pricing i finansijski model.
