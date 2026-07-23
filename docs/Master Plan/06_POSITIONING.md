# 06 — Pozicioniranje i ponuda vrednosti AgriX-a

**Status:** Review  
**Vlasnik:** osnivač AgriX-a  
**Poslednje ažuriranje:** 2026-07-23  
**Povezani dokumenti:** `02_STRATEGY.md`, `03_CUSTOMERS_AND_JOBS.md`, `04_MARKET.md`, `05_COMPETITION.md`, `07_PRODUCT_PORTFOLIO.md`, `08_PRODUCT_ROADMAP.md`, `10_PRICING_AND_PACKAGING.md`, `14_GO_TO_MARKET.md`, `15_SALES_PLAYBOOK.md`

---

## 1. Centralna proizvodna teza

AgriX nije desktop program kojem se podaci naknadno dostavljaju sa terena.

AgriX je distribuirani operativni sistem u kojem:

1. **otkupljivač/vagač na terenu** unosi otkup i izrađuje osnovna dokumenta na mestu nastanka događaja;
2. **vozač** evidentira zbirni transport, preuzimanje i status kretanja robe;
3. **PWA, GAS i sinhronizacioni sloj** prenose podatke u centralni poslovni sistem bez ponovnog prekucavanja;
4. **centralni desktop/backoffice** postaje kontrolni, dokumentni, finansijski i izveštajni centar;
5. **centralni operater** se primarno bavi kontrolom, prijemom, fakturama, SEF-om, bankom, saldima i izveštajima — ne ručnim prepisivanjem osnovnih događaja sa stanica.

> **Podatak se unosi jednom, na mestu nastanka, a zatim automatski puni centralnu bazu i ceo naredni dokumentacioni tok.**

Ovo je glavna vrednost AgriX-a i osnovna razlika u odnosu na sistem u kojem stanice rade na papiru, šalju Excel fajlove ili diktiraju podatke centrali.

---

## 2. Tržišna kategorija

### 2.1 Primarna kategorija

> **AgriX je terenski i centralni operativni sistem za organizovani otkup poljoprivrednih proizvoda.**

AgriX povezuje:

- otkupne stanice;
- otkupljivače/vagače;
- vozače i transport;
- kooperante;
- robu, klase, cene i ambalažu;
- osnovne dokumente nastale na terenu;
- prijem i centralnu bazu;
- fakture, SEF, banku i finansijsku kontrolu;
- management pregled i izveštavanje.

### 2.2 Proširena kategorija

> **AgriX je povezana platforma za organizovani otkup, upravljanje gazdinstvima i GGAP dokumentacioni tok.**

Za prvi kontakt sa hladnjačom ili organizovanim otkupljivačem primarna poruka ostaje: terenski unos, automatsko punjenje centrale i centralna kontrola.

### 2.3 Zašto ne „ERP za poljoprivredu“

AgriX nije univerzalno knjigovodstvo. Kupac može zadržati BizniSoft, PANTHEON ili drugi ERP, dok AgriX vodi duboki operativni tok od stanice do centrale.

Dozvoljena formulacija:

> AgriX je primarni operativni sistem otkupa, dok računovodstveni ERP može ostati knjigovodstveni sistem firme.

---

## 3. Centralna tržišna poruka

### 3.1 Jedna rečenica

> **Otkupci i vozači unose podatke i dokumente tamo gde posao nastaje, a AgriX automatski puni centralnu bazu za kontrolu, fakture i izveštaje.**

### 3.2 Kratka prodajna verzija

> AgriX omogućava da stanice, otkupljivači i vozači sami evidentiraju otkup i transport kroz PWA. Podaci automatski ulaze u centralni sistem, pa operater više ne prepisuje osnovna dokumenta, već kontroliše tok, radi prijem, fakture, finansije i izveštaje.

### 3.3 Verzija za vlasnika ili direktora

> Umesto da centrala čeka papire, telefonske izveštaje ili Excel fajlove sa stanica, AgriX daje managementu jedinstvenu sliku dok posao nastaje. Svaki otkup, dokument i transport ulazi u kontrolisani centralni tok.

### 3.4 Interna strateška formulacija

> AgriX pomera unos podataka iz centrale na mesto nastanka poslovnog događaja i pretvara centralnog operatera iz prepisivača u kontrolora poslovnog toka.

---

## 4. Osnovni poslovni model rada

### 4.1 Terenski otkupljivač/vagač

Otkupljivač:

- bira ili identifikuje kooperanta;
- bira stanicu, kulturu, klasu, cenu i ambalažu;
- evidentira bruto/neto i druge podatke otkupa;
- kreira osnovni otkupni dokument;
- štampa ili prosleđuje dokument prema podržanom toku;
- sinhronizuje podatak bez naknadnog unosa u centrali.

### 4.2 Vozač

Vozač:

- preuzima robu sa jedne ili više stanica;
- kreira i dopunjava zbirni transportni dokument;
- evidentira poslednju stanicu, količine i status;
- omogućava centrali da vidi šta je preuzeto, šta je u transportu i šta stiže na prijem.

### 4.3 Centralni operater

Centralni operater:

- kontroliše sinhronizovane podatke i izuzetke;
- završava prijem i dokumentne veze koje pripadaju centrali;
- radi fakture i SEF;
- obrađuje banku, avanse, salda i naloge za plaćanje;
- radi izveštaje i kontrole;
- rešava korekcije kroz kontrolisan storno/audit tok.

Centralni operater **ne treba rutinski da prepisuje otkupne podatke koje su otkupljivači već uneli na terenu**.

### 4.4 Management

Management dobija:

- pregled stanica i količina;
- status robe i transporta;
- dokumentacionu i finansijsku sliku;
- odstupanja i neusaglašenosti;
- pregled bez oslanjanja na telefonske pozive i ručne konsolidacije.

---

## 5. Primarni ICP

Najbolji kupac je firma koja:

- ima više stanica, lokacija ili terenskih tačaka;
- ima više otkupljivača/vagača i vozača;
- želi da osnovni dokumenti nastaju na terenu;
- danas prepisuje podatke u centrali ili spaja više fajlova i izvora;
- ima veći broj kooperanata i sezonski obim;
- ima prijem, ambalažu, logistiku, lager ili sledljivost;
- želi da centralni operater radi kontrolu, fakture i izveštaje umesto osnovnog unosa;
- prihvata standardan proizvod i konfiguraciju bez trajnog forka.

Prihod je pomoćni signal. Broj stanica, terenskih korisnika, dokumenata i ručnih prenosa podataka važniji su za procenu AgriX vrednosti.

---

## 6. Ponuda vrednosti po ulozi

### 6.1 Otkupljivač/vagač

- podatak se unosi jednom;
- dokument nastaje odmah na stanici;
- nema naknadnog diktiranja ili slanja papira centrali;
- sistem vodi obavezna polja i poslovna pravila;
- offline/sync tok čuva rad u realnim terenskim uslovima.

### 6.2 Vozač

- transport nastaje iz stvarnih terenskih podataka;
- zbirni dokument se vodi u toku preuzimanja;
- centrala vidi robu koja dolazi;
- smanjuje se telefonsko koordinisanje i naknadno rekonstruisanje rute.

### 6.3 Centralna administracija i finansije

- nema rutinskog ponovnog unosa osnovnih otkupa;
- koristi već nastale podatke za prijem, fakture i finansije;
- kontroliše izuzetke umesto da ručno gradi celu bazu;
- dobija povezan dokumentni lanac, storno i audit;
- lakše prenosi posao na obučenu zamenu.

### 6.4 Vlasnik i management

- vidi podatke dok posao nastaje;
- dobija kontrolu mreže bez mikromenadžmenta;
- ranije otkriva odstupanja;
- može da poveća broj stanica bez proporcionalnog rasta centralnog administrativnog rada.

---

## 7. Pozicioniranje prema alternativama

### 7.1 Papir, telefon i Excel fajlovi

> AgriX uklanja ručni prenos podataka između terena i centrale. Osnovni događaj nastaje digitalno na stanici i automatski ulazi u centralni sistem.

### 7.2 Desktop-only program

> Desktop-only sistem centralizuje unos, ali ne rešava mesto nastanka podatka. AgriX omogućava da otkupljivači i vozači sami formiraju operativne podatke i dokumente, dok centrala zadržava kontrolu.

### 7.3 Generički ERP

> ERP vodi knjigovodstvo i opšte poslovne funkcije; AgriX vodi terenski i centralni tok organizovanog otkupa i prosleđuje pripremljene podatke prema računovodstvu.

### 7.4 Uski program za otkup

> Otkupni list nije izolovan dokument. U AgriX-u je početak lanca koji nastavljaju transport, prijem, fakture, finansije, lager i management kontrola.

### 7.5 Infosys i drugi incumbent sistemi

Dve postojeće migracije potvrđuju replacement market, ali ne dokazuju opštu superiornost. Do završetka win intervjua dozvoljeno je tvrditi samo da AgriX ima praktično iskustvo sa migracijom iz specijalizovanog sistema.

---

## 8. Trenutna proizvodna ponuda

### 8.1 Centralna produkciona celina

AgriX Enterprise obuhvata povezani sistem:

- PWA Otkupac;
- PWA Vozač;
- sinhronizacioni i transportni sloj;
- desktop centralni backoffice;
- Management PWA;
- dokumentni, finansijski i izveštajni tok.

PWA Otkupac i PWA Vozač nisu periferni dodaci. Oni su glavni izvori terenskih poslovnih događaja.

### 8.2 Odvojeno procenjivati

Od same PWA aplikacije odvojeno se procenjuju:

- standardizovan kiosk režim;
- konkretan tablet paket;
- termalna štampa i printer bridge;
- remote device management;
- hardverska garancija i zamena.

Nedovoljno zreo hardverski paket ne sme automatski da spusti status dobro funkcionalne PWA aplikacije.

### 8.3 Budući proizvodi

- Gazdinstvo zahteva activation, retention i willingness-to-pay dokaz;
- GGAP ostaje discovery/pilot dok ne postoji stručni domain owner i validiran sadržaj.

---

## 9. Dozvoljene marketinške tvrdnje

Dozvoljeno, kada je scope tačan:

- „Osnovni otkupni podaci i dokumenti nastaju na terenu.“
- „Otkupci i vozači sami unose događaje koji automatski pune centralnu bazu.“
- „Centralni operater se fokusira na kontrolu, fakture, finansije i izveštaje.“
- „AgriX povezuje PWA terenski rad sa centralnim desktop backoffice-om.“
- „Podatak se ne prepisuje ponovo ako je već pravilno unet na mestu nastanka.“
- „AgriX može raditi uz postojeći računovodstveni ERP.“
- „Tri firme koriste AgriX, a dva klijenta su prešla sa Infosys sistema.“

Brojčane tvrdnje moraju imati datum, izvor i scope.

---

## 10. Privremeno nedozvoljene tvrdnje

Bez dodatnog dokaza ne koristiti:

- „eliminiše sve greške“;
- „nikada ne gubi podatke“;
- „radi bez interneta u svakoj funkciji“;
- „podržava svaki printer i svaki uređaj“;
- „svaka migracija je laka“;
- „potpuna zamena ERP-a“;
- „spreman za neograničen broj stanica“;
- „potpun GGAP compliance“ ili garantovana sertifikacija.

---

## 11. Demo struktura

Demo mora pratiti stvarni tok, ne spisak ekrana:

1. otkupljivač na stanici kreira otkup;
2. dokument i podatak ulaze u sinhronizacioni tok;
3. vozač preuzima robu i kreira zbirni transport;
4. centrala dobija podatke bez ponovnog unosa;
5. operater završava prijem, fakturu i finansijski tok;
6. management vidi rezultat i odstupanja.

Glavna demo poruka:

> **Od prvog unosa na stanici do fakture i izveštaja u centrali — bez ponovnog prekucavanja istog poslovnog događaja.**

---

## 12. Dokazi koje treba meriti

- procenat otkupa unetih direktno na terenu;
- broj osnovnih dokumenata nastalih bez centralnog ručnog unosa;
- broj ponovnih unosa eliminisanih po otkupu;
- vreme od terenskog događaja do vidljivosti u centrali;
- sync success rate;
- broj duplikata i ručnih korekcija;
- vreme centralnog operatera po 100 otkupa;
- odnos vremena utrošenog na osnovni unos naspram kontrole/faktura/izveštaja;
- broj stanica koje jedan centralni operater može da podrži;
- support minuti po stanici.

---

## 13. Odluke

1. PWA Otkupac i PWA Vozač su centralni operativni kanali AgriX Enterprise-a.
2. Desktop je centralni backoffice i canonical sloj nakon sinhronizacije, ali nije zamišljen kao mesto rutinskog prepisivanja terenskih događaja.
3. Osnovna vrednost proizvoda je automatsko punjenje centralne baze iz dokumenata nastalih na terenu.
4. Centralni operater treba da se bavi kontrolom, prijemom, fakturama, finansijama i izveštajima.
5. Kiosk, tablet i termalna štampa imaju odvojene readiness statuse od same PWA aplikacije.
6. Enterprise ponuda mora biti demonstrirana kao jedan end-to-end tok: teren → sinhronizacija → centrala → faktura/izveštaj.
7. Pricing mora vrednovati broj stanica i terenskih tokova, ne samo desktop licencu.
8. Roadmap mora davati PWA terenskom toku najmanje isti strateški značaj kao centralnom backoffice-u.
9. Gazdinstvo i GGAP ostaju zasebni proizvodi sa posebnim dokaznim pragovima.
10. Pozicioniranje se dopunjava merljivim rezultatima i win intervjuima, ali centralna PWA teza nije hipoteza već namera proizvoda i postojeći operativni model.
