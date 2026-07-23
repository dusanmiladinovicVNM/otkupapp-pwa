# 07 — Portfolio proizvoda AgriX-a

**Status:** Review  
**Vlasnik:** osnivač AgriX-a  
**Poslednje ažuriranje:** 2026-07-23  
**Povezani dokumenti:** `02_STRATEGY.md`, `02A_GGAP_STRATEGY.md`, `03_CUSTOMERS_AND_JOBS.md`, `06_POSITIONING.md`, `08_PRODUCT_ROADMAP.md`, `10_PRICING_AND_PACKAGING.md`, `14_GO_TO_MARKET.md`, `15_SALES_PLAYBOOK.md`, `16_ONBOARDING_AND_IMPLEMENTATION.md`

---

## 1. Svrha

Ovo poglavlje definiše:

- šta je AgriX proizvod, a šta tehnička komponenta, usluga ili hardver;
- koji proizvodni stubovi postoje;
- šta se danas može prodavati kao standardna produkciona ponuda;
- šta je opciona ekstenzija;
- šta se nudi samo kao kontrolisani pilot;
- šta je planirano i ne sme se predstavljati kao završeno;
- koje funkcije moraju ostati zajedno da bi kupac dobio celovit poslovni ishod;
- gde prestaje AgriX, a počinju računovodstveni ERP, banka, SEF, konsultant, sertifikaciono telo ili hardverski dobavljač;
- kako se sprečava da portfolio postane skup klijentskih forkova i nepovezanih funkcija.

Portfolio nije lista svih ekrana, VBA modula, tabela i API endpointa. Kupac ne kupuje tehničku arhitekturu. Kupuje kontrolisan poslovni tok, definisanu odgovornost, implementaciju i podršku.

---

## 2. Kritično pravilo: tri različita statusa

Postojanje funkcije u kodu nije isto što i dokaz da je spremna za standardnu prodaju.

Za svaku proizvodnu celinu vode se tri odvojene procene:

### 2.1 Implementacioni status

- `Implemented` — funkcionalnost postoji end-to-end u glavnom codebase-u;
- `Partial` — postoji samo deo toka ili nedostaju kritične integracije, UX ili kontrole;
- `Planned` — funkcionalnost je definisana, ali nije završena;
- `Deprecated` — ne razvija se dalje i ne nudi se novim klijentima.

### 2.2 Dokazni status

- `Production-proven` — koristi se u realnom radu i prošla je relevantan sezonski ili poslovni ciklus;
- `Limited production evidence` — koristi se, ali u malom broju firmi, ograničenom obimu ili bez pune sezonske potvrde;
- `Pilot evidence` — testirana je sa kontrolisanim korisnicima i unapred definisanim scope-om;
- `Unvalidated` — tehnički postoji ili je planirana, ali nema dovoljno poslovnog dokaza.

### 2.3 Komercijalni status

- `Standard offer` — može se nuditi svakom kvalifikovanom klijentu uz standardan onboarding;
- `Optional extension` — nudi se kada klijent ima relevantan proces i prihvata dodatni scope;
- `Pilot only` — nudi se samo imenovanom pilot-klijentu uz success criteria i ograničenja;
- `Not for sale` — ne sme biti u ugovoru ili ponudi kao završena produkciona funkcionalnost.

`DECISION`: demo, ponuda i ugovor moraju koristiti komercijalni status, a ne samo činjenicu da određeni kod postoji.

---

## 3. Arhitektura portfolija

AgriX ima tri puna proizvodna stuba i jedan zajednički platform layer:

1. **AgriX Enterprise** — operativni sistem firme koja organizuje otkup;
2. **AgriX Gazdinstvo** — farm-management proizvod kooperanta/proizvođača;
3. **AgriX GGAP** — dokumentacioni, dokazni i compliance workflow;
4. **AgriX Platform Services** — zajednički tehnički i operativni sloj koji omogućava pouzdan rad sva tri proizvoda.

Terenski PWA tokovi, Management PWA, logistika, finansijske integracije i hardver nisu zasebne kompanije niti nezavisni proizvodi bez konteksta. Oni su moduli, kanali korišćenja ili delivery komponente navedenih proizvodnih stubova.

---

## 4. AgriX Platform Services

Platform Services se ne prodaje kao zasebna licenca. Njegov trošak i vrednost moraju biti ugrađeni u cenu proizvoda.

### 4.1 Obuhvat

- zajednički codebase bez trajnih klijentskih forkova;
- konfiguracija firme, stanica, korisnika, uloga i feature flags;
- instalacija i setup radne stanice;
- MasterSync i kontrolisane projekcije prema Google Sheets/PWA sloju;
- autentifikacija i role/entity autorizacija;
- offline queue, idempotency i kontrolisana sinhronizacija;
- transakcioni wrapper-i i rollback za desktop-owned podatke;
- audit i storno tokovi;
- self-update i upravljanje verzijama;
- monitoring, error logging i health signali;
- backup, journaling i recovery pomoć;
- Production Health Check i release gates;
- konfiguracioni import/export i standardizovan onboarding;
- podrška i incident dijagnostika.

### 4.2 Status

| Dimenzija | Status |
|---|---|
| Implementacija | `Implemented` |
| Dokaz | `Production-proven`, uz ograničen broj klijenata |
| Komercijalno | sastavni deo svih plaćenih proizvoda |

### 4.3 Pravilo

Platform Services nije „besplatna pozadina“. Monitoring, update, backup, recovery i podrška imaju realan operativni trošak i moraju biti uključeni u unit economics i godišnju cenu.

---

## 5. AgriX Enterprise

### 5.1 Uloga proizvoda

AgriX Enterprise je primarni komercijalni proizvod i operativni sistem firme.

Njegova uloga je da poveže:

- centralu;
- otkupne stanice;
- kooperante i dobavljače;
- robu i klase;
- dokumentni lanac;
- ambalažu i repromaterijal;
- logistiku i vozila;
- prijem, lager i sledljivost;
- kupce i izlaz robe;
- finansijske tokove;
- management kontrolu.

Enterprise ne mora da zameni knjigovodstveni ERP. Standardna arhitektura je da AgriX bude primarni operativni sistem, dok BizniSoft, PANTHEON ili drugi računovodstveni sistem ostaje knjigovodstveni source of truth tamo gde je to racionalno.

---

## 6. Enterprise Core

Enterprise Core je minimalna koherentna celina koju novi standardni klijent kupuje.

### 6.1 Funkcionalni obuhvat

#### Master podaci i priprema sezone

- firme i konfiguracija;
- kooperanti/dobavljači;
- otkupne stanice;
- korisnici i uloge;
- kulture, artikli, klase i cenovnici;
- parcele i povezani osnovni podaci gde se koriste;
- kupci, vozila, vozači i drugi relevantni registri;
- priprema nove sezone bez kopiranja klijentskog forka.

#### Otkup i dokumentni lanac

- unos i import otkupa;
- bruto/neto i klasni tokovi;
- otkupni listovi i povezani dokumenti;
- otpremnica;
- zbirna otpremnica;
- prijemnica;
- faktura i stavke;
- kontrolisani brojevi dokumenata;
- storno, korekcije i audit trag;
- integritet veza između dokumenata.

#### Ambalaža i osnovna zaduženja

- izdata i vraćena ambalaža;
- revers i kartica ambalaže;
- povezivanje ambalaže sa kooperantom i otkupnim dokumentima;
- osnovna kontrola otvorenih zaduženja.

#### Osnovni prijem, lager i izveštavanje

- prijem robe;
- osnovne količinske kontrole;
- pregled toka robe po relevantnim dimenzijama;
- standardni operativni i management izveštaji;
- exporti potrebni za dalji računovodstveni ili kontrolni tok.

#### Management Control

- Management PWA read-modeli i dashboardi koji su aktivni za konkretnog klijenta;
- pregled stanica, količina, dokumenata, kartica i odstupanja;
- pristup u skladu sa ulogama;
- monitoring i operativni signali.

### 6.2 Status

| Dimenzija | Status |
|---|---|
| Implementacija | `Implemented` |
| Dokaz | `Production-proven` kod postojećih klijenata |
| Komercijalno | `Standard offer` |

### 6.3 Obavezna granica

Enterprise Core ne sme biti razbijen na deset mikro-doplata tako da osnovni proizvod više ne završava kompletan posao. Cenovni paketi mogu ograničiti obim, broj firmi, stanica, korisnika ili napredne module, ali core dokumentni integritet, monitoring i sigurnosne kontrole nisu opciona doplata.

---

## 7. Enterprise ekstenzije

Ekstenzije su funkcionalno povezane sa Enterprise Core-om, ali nisu potrebne svakom klijentu i stvaraju poseban implementation/support cost.

### 7.1 Finance & Regulatory Integration

#### Obuhvat

- SEF workflow i statusi;
- slanje i praćenje relevantnih dokumenata;
- bankarski izvod i BankaImport staging;
- mapiranje partnera i transakcija;
- rasknjižavanje uplata i isplata;
- avansi, salda i status naplate;
- priprema naloga za plaćanje;
- finansijske kartice i kontrolni izveštaji;
- integracioni export/import prema postojećem računovodstvu.

#### Status

| Dimenzija | Status |
|---|---|
| Implementacija | `Implemented` za postojeće podržane tokove |
| Dokaz | `Limited production evidence`; obim zavisi od klijenta i banke/ERP-a |
| Komercijalno | `Optional extension` |

#### Granica

AgriX ne postaje univerzalni knjigovodstveni program. Svaka nova banka, ERP ili poseban format zahteva procenu ponovne upotrebljivosti, održavanja i odgovornosti.

---

### 7.2 Logistics & Fleet

#### Obuhvat

- Dispečer;
- potražnja/preuzimanje robe sa stanica;
- raspodela vozila i vozača;
- rute, statusi i kapaciteti;
- zbirne otpremnice i poslednja stanica;
- Vozač PWA tok;
- pregled realizovanih i otvorenih transporta;
- management pregled logistike.

#### Status

| Dimenzija | Status |
|---|---|
| Implementacija | `Implemented` za dokumentovane tokove |
| Dokaz | `Limited production evidence` |
| Komercijalno | `Optional extension`; standardna prodaja tek posle klijentskog fit-checka |

#### Granica

Ovo nije generički TMS za sve vrste transporta. Modul je namenjen logistici koja neposredno prati organizovani otkup i kretanje robe između stanica, vozila i prijema.

---

### 7.3 Inputs, Agrohemija & Cooperant Balance

#### Obuhvat

- repromaterijal i agrohemija;
- izdavanje i razduženje;
- zaduženje kooperanta;
- kartice, saldo i povezivanje sa proizvodnjom/otkupom;
- management izdavanje artikala;
- artikli, doziranje i pomoćni kontrolni tokovi;
- otpremnice i dokaz izdavanja;
- povezivanje sa Gazdinstvom gde je aktivirano.

#### Status

| Dimenzija | Status |
|---|---|
| Implementacija | `Implemented` za postojeći scope |
| Dokaz | `Limited production evidence` |
| Komercijalno | `Optional extension` ili deo višeg Enterprise paketa |

#### Granica

Modul ne sme da se prodaje kao agronomska garancija. Preporuka doziranja i upozorenja moraju imati jasne izvore, pravila i odgovornost korisnika/stručnog lica.

---

### 7.4 Advanced Warehouse, Pallets & Traceability

#### Obuhvat

- napredniji magacinski tok;
- palete i jedinice skladištenja;
- povezivanje prijema, prerade, lagera i izlaza;
- sledljivost robe i dokumenata;
- povezivanje kooperanta/parcele sa prijemom i kupcem kada podaci postoje;
- kontrolni i izvozni izveštaji;
- podrška budućem GGAP dokaznom toku.

#### Status

| Dimenzija | Status |
|---|---|
| Implementacija | `Implemented/Partial`, zavisno od poddomena |
| Dokaz | `Limited production evidence` |
| Komercijalno | `Optional extension`; scope se potvrđuje u discovery-ju |

#### Granica

Ne obećavati generički WMS, proizvodno planiranje ili pun MES bez posebnog dokaza i roadmap odluke.

---

## 8. Enterprise Field Operations

Enterprise Field Operations obuhvata rad na mestu otkupa i nije zaseban proizvod odvojen od Enterprise-a.

### 8.1 PWA Otkupac

#### Cilj

- brz unos otkupa na mestu nastanka;
- rad pri slaboj ili nestabilnoj vezi;
- lokalni queue i kasnija sinhronizacija;
- identifikacija kooperanta;
- kontrola cene, klase, ambalaže i dokumenta;
- eliminacija kasnijeg prepisivanja;
- povezivanje sa centralnim dokumentnim tokom.

#### Status

| Dimenzija | Status |
|---|---|
| Implementacija | `Implemented/Partial` prema aktivnom release scope-u |
| Dokaz | `Pilot evidence` ili pre-production evidence |
| Komercijalno | `Pilot only` dok ne prođe release i sezonske kriterijume |

### 8.2 Kiosk tablet i termalna štampa

#### Obuhvat

- standardizovan Android/kiosk uređaj;
- zaključan operativni režim;
- pouzdana termalna štampa na stanici;
- printer bridge ili druga kontrolisana integracija;
- remote konfiguracija i podrška;
- rezervni uređaj i plan zamene.

#### Status

| Dimenzija | Status |
|---|---|
| Implementacija | `Partial/Planned` kao standardizovan delivery paket |
| Dokaz | `Unvalidated` u planiranom obimu |
| Komercijalno | `Pilot only`; hardver se iskazuje odvojeno |

### 8.3 Production gate za Field Operations

Pre statusa `Standard offer` moraju biti potvrđeni:

1. kompletan otkupni tok bez desktop ručnog popravljanja;
2. offline unos, retry, idempotency i recovery;
3. stabilna štampa na odobrenom hardveru;
4. kontrola duplikata i pogrešnih dokumenata;
5. jasan sync status za vagača;
6. remote support procedura;
7. najmanje jedan realan sezonski pilot;
8. merljivo vreme unosa i stopa grešaka;
9. rollback/recovery procedura;
10. standardizovan onboarding stanice.

`DECISION`: PWA Otkupac, kiosk i termalna štampa ne smeju se u ponudi predstavljati kao standardna produkciona komponenta pre prolaska ovog gate-a.

---

## 9. Management PWA

Management PWA je kanal korišćenja Enterprise proizvoda, ne zasebna firma niti samostalni BI proizvod.

### 9.1 Obuhvat

- pregled operativnih rezultata;
- stanice, količine i statusi;
- kartice i relevantni saldi;
- dispatch/logistički pregled gde je modul aktivan;
- upravljanje pojedinim kontrolisanim operacijama;
- QR identifikacija kooperanta;
- izdavanje repromaterijala/agrohemije gde je aktivirano;
- izveštaji, monitoring i upozorenja;
- pristup prilagođen management ulozi.

### 9.2 Status

| Dimenzija | Status |
|---|---|
| Implementacija | `Implemented` za postojeći scope |
| Dokaz | `Production-proven` kod postojećih klijenata |
| Komercijalno | deo `Enterprise Core`; napredne operacije zavise od modula |

### 9.3 Granica

Management PWA nije zamena za desktop source of truth. To je kontrolisani mobilni/web pogled i operativni interfejs nad podacima koje poseduje odgovarajući canonical sloj.

---

## 10. AgriX Gazdinstvo

### 10.1 Uloga proizvoda

AgriX Gazdinstvo je pun farm-management proizvod za proizvođača, ali istovremeno i B2B2C produžetak odnosa hladnjače sa kooperantima.

### 10.2 Planirani funkcionalni obuhvat

- kartica i saldo prema hladnjači;
- zaduženje i razduženje repromaterijala;
- parcele, kulture, površine i GIS;
- prognoza i upozorenja po parceli;
- tretmani, mere, doze i karenca;
- oprema, vreme rada i lokacija;
- lager agrohemije;
- troškovi po kategoriji i parceli;
- proizvodnja i knjiga polja;
- sezonski bilans ukupno i po parceli;
- dokumenti, fotografije i fiskalni računi;
- offline-first rad i sinhronizacija;
- buduća veza sa GGAP evidencijama.

### 10.3 Varijante proizvoda

Nazivi `Partner`, `Basic` i `Pro` predstavljaju radni packaging model. Tačna granica funkcija i cena zaključava se u `10_PRICING_AND_PACKAGING.md` tek nakon activation, retention i willingness-to-pay testa.

#### Partner

Nalog koji hladnjača distribuira kooperantu radi povezivanja sa Enterprise tokom. Mora imati dovoljno vrednosti da se aktivira i koristi, ali ne sme besplatno sadržati ceo Pro proizvod.

#### Basic

Samostalni osnovni farm-management paket za vođenje parcela, tretmana, troškova i sezonskog pregleda.

#### Pro

Napredniji paket sa naprednim analizama, automatizacijom, Digitalnim agronomom, pametnim doziranjem, širim upozorenjima i drugim premium funkcijama koje prođu validaciju.

### 10.4 Status

| Dimenzija | Status |
|---|---|
| Implementacija | `Implemented/Partial` po funkcionalnim domenima |
| Dokaz | `Pilot evidence`; komercijalna retencija i WTP nisu potvrđeni |
| Komercijalno | `Pilot only` / kontrolisana rana ponuda |

### 10.5 Production gate

Pre skalirane prodaje potrebno je potvrditi:

- activation funnel;
- procenat aktivnih kooperanata nakon 30, 90 i 180 dana;
- retenciju kroz sezonu i između sezona;
- willingness-to-pay za Basic i Pro;
- support cost po aktivnom korisniku;
- kvalitet offline sync-a;
- tačnost GIS/meteo/agronomskih funkcija;
- jasnu odgovornost za preporuke;
- vrednost za hladnjaču kao distributera;
- da Partner nalog povećava, a ne smanjuje, disciplinu podataka.

---

## 11. AgriX GGAP

### 11.1 Uloga proizvoda

AgriX GGAP je treći puni proizvodni stub koji koristi podatke iz Enterprise i Gazdinstvo sistema za liste, evidencije, dokaze, zadatke, neusaglašenosti i audit readiness.

### 11.2 Status

| Dimenzija | Status |
|---|---|
| Implementacija | `Planned` / discovery |
| Dokaz | `Unvalidated` |
| Komercijalno | `Not for sale` kao produkcioni proizvod |

### 11.3 Dozvoljena rana ponuda

Dozvoljeni su samo:

- discovery sa kvalifikovanim klijentom;
- mapiranje postojećih listi i dokumentacije;
- ograničeni prototip;
- imenovani pilot sa stručnim domain owner-om;
- ugovor koji jasno navodi pilot scope i ne garantuje sertifikaciju.

### 11.4 Zabranjene tvrdnje

AgriX GGAP se ne predstavlja kao:

- sertifikaciono telo;
- nezavisni auditor;
- garancija dobijanja ili zadržavanja sertifikata;
- zamena za GGAP konsultanta;
- automatski produkcioni modul pre stručne validacije standarda i sadržaja.

---

## 12. Usluge

Usluge nisu trajni klijentski fork. One omogućavaju implementaciju standardnog proizvoda.

### 12.1 Standardne usluge

- discovery i procesno mapiranje;
- konfiguracija firme;
- migracija master podataka;
- migracija otvorenih stanja i istorije kada je ugovoreno;
- onboarding i obuka;
- priprema radnih stanica;
- povezivanje podržanih PWA/GAS/Sheets komponenti;
- konfiguracija standardnih integracija;
- go-live podrška;
- sezonski readiness review;
- periodični health check.

### 12.2 Posebne usluge

- kompleksna migracija sa Infosys-a ili drugog incumbent sistema;
- nova banka ili ERP integracija;
- posebni izveštaji koji imaju jasnu mogućnost ponovne upotrebe;
- procesni consulting u okviru AgriX domena;
- hardverska instalacija i remote management;
- custom import/export adapter.

### 12.3 Pravilo za posebne zahteve

Svaki zahtev prolazi kroz četiri pitanja:

1. Da li rešava problem primarnog ICP-a?
2. Da li može postati konfigurabilna funkcija zajedničkog proizvoda?
3. Da li prihod pokriva razvoj, testiranje, dokumentaciju i buduće održavanje?
4. Da li ugrožava sezonski roadmap ili pouzdanost core proizvoda?

Ako odgovor nije dovoljno pozitivan, zahtev se odbija, odlaže ili rešava izvan glavnog proizvoda bez trajnog forka.

---

## 13. Hardver

Hardver je delivery i profitna komponenta, ali nije osnovni softverski proizvod.

### 13.1 Potencijalni portfolio

- standardizovan kiosk tablet;
- termalni štampač;
- rezervni uređaj;
- stalak, zaštita i napajanje;
- konfiguracija i kiosk zaključavanje;
- remote device management;
- zamena i garancijski tok;
- opciona vaga/printer integracija kada je podržana.

### 13.2 Komercijalno pravilo

Hardver se u ponudi prikazuje odvojeno:

- nabavna vrednost;
- prodajna cena ili najam;
- konfiguracija;
- instalacija;
- garancija;
- rezervna oprema;
- održavanje i zamena.

Prihod od hardvera se ne prikazuje kao softverski recurring prihod, a marža se računa tek nakon rada, reklamacija, zalihe i rizika.

---

## 14. Trenutno dozvoljena prodajna ponuda

### 14.1 Standardno prodavati

**AgriX Enterprise Core**, uključujući:

- desktop operativni backbone;
- master podatke i pripremu sezone;
- otkup i dokumentni lanac;
- osnovnu ambalažu/reverse;
- osnovni prijem i izveštavanje;
- Management PWA u aktivnom podržanom scope-u;
- monitoring, update, backup i support layer;
- standardan onboarding i konfiguraciju.

### 14.2 Prodavati kao opcionu ekstenziju

Nakon discovery-ja i tehničke potvrde:

- Finance & Regulatory Integration;
- Logistics & Fleet;
- Inputs, Agrohemija & Cooperant Balance;
- Advanced Warehouse, Pallets & Traceability;
- kompleksna migracija i posebni adapteri.

### 14.3 Nuditi samo kao pilot

- PWA Otkupac kao standardni terenski production kanal;
- kiosk tablet paket;
- termalna štampa na stanici;
- Gazdinstvo Partner/Basic/Pro u skaliranom komercijalnom modelu;
- nove ili nepotvrđene integracije.

### 14.4 Ne prodavati kao završen proizvod

- AgriX GGAP;
- generički ERP;
- generički TMS/WMS/MES;
- garantovanu sertifikaciju;
- neograničen custom development;
- funkcionalnost koja postoji samo u demonstracionom ili nedovršenom toku.

---

## 15. Zavisanosti između proizvoda i modula

| Celina | Obavezna zavisnost | Napomena |
|---|---|---|
| Enterprise Core | Platform Services | ne može se isključiti monitoring, update i integrity layer |
| Management PWA | Enterprise Core + projekcije podataka | nije zaseban source of truth |
| Finance & Regulatory | Enterprise Core + klijentska konfiguracija | zavisi od banke, ERP-a i SEF pravila |
| Logistics & Fleet | Enterprise Core + stanice/vozila/vozači | namenjeno logistici otkupa |
| Field Operations | Enterprise Core + GAS/Sheets/PWA + odobren hardver | pilot dok ne prođe sezonski gate |
| Gazdinstvo Partner | Platform + veza sa Enterprise firmom | mora imati jasan tenant/entity model |
| Gazdinstvo Basic/Pro | Platform; Enterprise veza opciona | samostalni farm-management proizvod |
| GGAP Enterprise | Enterprise podaci + stručni sadržaj | ručni fallback ne sme poništiti glavnu vrednost automatizacije |
| GGAP Gazdinstvo | Gazdinstvo podaci + stručni sadržaj | ne garantuje sertifikaciju |

---

## 16. Portfolio pravila za demo, ponudu i ugovor

1. Demo mora jasno označiti `production`, `optional`, `pilot` i `planned` funkcije.
2. Planirana funkcija ne ulazi u ugovor kao postojeća osim ako ima zaseban milestone, cenu, acceptance criteria i pravo na odlaganje/raskid.
3. Klijent ne dobija poseban fork; dobija konfiguraciju, feature flags i eventualno novu reusable funkciju.
4. Svaki modul ima definisan owner, support boundary i release gate.
5. Integracija sa eksternim sistemom mora imati tačno naveden format, odgovornost i fallback.
6. Hardver i usluge iskazuju se odvojeno od recurring softverske licence.
7. „Neograničena podrška“ se ne obećava bez definisanog SLA i fair-use granice.
8. Pilot ne postaje trajno besplatan production sistem.
9. Postojanje funkcije kod jednog klijenta ne znači da je automatski deo svakog paketa.
10. Core sigurnost, audit, backup i monitoring ne smeju se uklanjati radi niže cene.

---

## 17. Readiness gate za prelazak u Standard offer

Proizvodna celina prelazi u `Standard offer` tek kada ima:

- jasno definisan problem i ICP;
- end-to-end tok bez ručnih skrivenih koraka osnivača;
- release i regression kriterijume;
- najmanje jedan relevantan produkcioni dokaz, a za sezonske tokove najmanje jedan realan sezonski ciklus;
- dokumentovan onboarding;
- monitoring i incident dijagnostiku;
- backup/recovery ili jasan failure model;
- standardnu konfiguraciju bez forka;
- support dokumentaciju;
- poznat support cost;
- cenu ili packaging pravilo;
- ugovornu granicu odgovornosti;
- owner-a i roadmap održavanja;
- merljive success KPI-jeve.

---

## 18. Portfolio KPI-jevi

### Enterprise

- broj aktivnih firmi;
- broj aktivnih stanica;
- broj obrađenih otkupa i dokumenata;
- onboarding sati po firmi;
- support sati po firmi;
- broj P0/P1 incidenata u sezoni;
- procenat standardne konfiguracije bez custom koda;
- renewal i churn;
- ARR po firmi i stanici;
- usage po opcionom modulu.

### Field Operations

- prosečno vreme unosa otkupa;
- procenat offline unosa;
- sync success rate;
- duplicate/error rate;
- print success rate;
- broj intervencija po uređaju;
- vreme osposobljavanja nove stanice.

### Gazdinstvo

- activation rate;
- 30/90/180-day active rate;
- retencija između sezona;
- aktivne parcele i tretmani;
- procenat Partner→Basic/Pro konverzije;
- ARPU i support cost;
- broj ručnih duplih unosa eliminisanih povezivanjem sa Enterprise-om.

### GGAP

- broj pilot firmi/gazdinstava;
- procenat automatski izvedenih podataka;
- vreme pripreme dokumentacije;
- kompletiranost dokaza;
- otvorene/zatvorene neusaglašenosti;
- support i domain-maintenance cost.

---

## 19. Ključni rizici

| Rizik | Uticaj | Zaštita |
|---|---|---|
| Kupcu se obeća funkcija koja postoji samo u kodu | visok | tri statusa i obavezna oznaka u ponudi |
| Enterprise postane zbir nepovezanih doplata | visok | koherentan Core i ograničen broj ekstenzija |
| Svaki klijent stvara novi fork | kritičan | konfiguracija, reusable kriterijum i architecture review |
| Pilot ostane trajno neplaćen i bez exit kriterijuma | visok | pilot ugovor, rok, KPI i odluka scale/stop |
| Gazdinstvo prerano optereti support | visok | activation/retention gate i kontrolisana distribucija |
| GGAP se proda pre stručne validacije | kritičan | `Not for sale`, domain owner i ograničen pilot |
| Hardver pojede maržu i vreme | visok | odvojena ekonomika, standardni modeli i rezervna oprema |
| Integracije postanu neodrživi custom adapteri | visok | podržani formati, verzionisanje i cena održavanja |
| Previše modula zamagli centralnu poruku | srednji | Enterprise Core kao glavni prodajni narativ |

---

## 20. Otvorene odluke za pricing i roadmap

Sledeća poglavlja moraju zaključiti:

1. tačan osnovni Enterprise paket i njegove limite;
2. da li se cena primarno vezuje za firmu, broj stanica, obim ili kombinaciju;
3. koji moduli su bundle, a koji posebna ekstenzija;
4. standardni scope Management PWA;
5. komercijalni model Finance & Regulatory modula;
6. komercijalni model Logistics & Fleet modula;
7. granice Inputs/Agrohemija i Advanced Warehouse modula;
8. uslove i cenu Infosys migracije;
9. Partner/Basic/Pro granice Gazdinstva;
10. pilot cenu i success criteria za Field Operations;
11. hardver: prodaja, najam ili obe opcije;
12. minimalni godišnji prihod koji opravdava support i sezonski rizik.

---

## 21. Predložene portfolio odluke

### POR-001 — Tri proizvoda, jedna platforma

AgriX ima tri puna proizvoda: Enterprise, Gazdinstvo i GGAP. Platform Services je zajednički obavezni sloj.

### POR-002 — Enterprise kao komercijalno jezgro

Enterprise Core ostaje primarni proizvod, glavni izvor prihoda i početna tačka prodaje.

### POR-003 — Management PWA je deo Enterprise-a

Management PWA se ne prodaje kao nepovezan dashboard, već kao kontrolni kanal Enterprise proizvoda.

### POR-004 — Field Operations je Enterprise modul

PWA Otkupac, kiosk i termalna štampa čine Field Operations ekstenziju i ostaju `Pilot only` do sezonske validacije.

### POR-005 — Ograničen broj komercijalnih ekstenzija

Funkcionalnosti se grupišu u poslovno koherentne ekstenzije; ne uvodi se mikro-naplata svakog ekrana i izveštaja.

### POR-006 — Bez klijentskih forkova

Poseban zahtev ulazi u proizvod samo kada je konfigurabilan i ponovljivo vredan ciljnom tržištu.

### POR-007 — Gazdinstvo zahteva tržišni dokaz

Partner, Basic i Pro ostaju radni packaging model dok activation, retention, willingness-to-pay i support cost ne budu izmereni.

### POR-008 — GGAP nije još produkciona ponuda

GGAP se razvija kroz stručni discovery i ograničen pilot; ne prodaje se kao završena compliance garancija.

### POR-009 — Hardver i usluge imaju odvojenu ekonomiku

Hardver, implementacija, migracija i posebne integracije prikazuju se odvojeno od recurring softverskog prihoda.

### POR-010 — Komercijalni status upravlja obećanjem

Ponuda i ugovor smeju obećati samo ono što ima odgovarajući komercijalni status i acceptance granicu.

---

## 22. Sledeći korak

`08_PRODUCT_ROADMAP.md` treba da prevede ovaj portfolio u vremenski redosled:

1. šta mora da bude production-ready pre naredne sezone;
2. šta podiže Enterprise readiness i onboarding kapacitet;
3. šta je Field Operations pilot scope;
4. šta se odlaže posle sezone;
5. koliko kapaciteta ostaje za Gazdinstvo i GGAP discovery;
6. koji moduli ne smeju napredovati dok core pouzdanost ili support nisu spremni.

Tek nakon roadmap odluke `10_PRICING_AND_PACKAGING.md` može precizno definisati pakete i cenu bez prodavanja nedovršenog scope-a.