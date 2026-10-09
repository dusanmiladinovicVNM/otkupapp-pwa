# AgriX — funkcionalne bele zone i proizvodni roadmap

> **Svrha:** product-level pogled na ono što AgriX funkcionalno još nema ili nema kao zaokružen poslovni capability.
> Ovo **nije code review** i ne govori o kvalitetu implementacije, arhitekturi ili refaktoru.
> Fokus je: šta nedostaje da AgriX pređe iz veoma jakog sistema za otkup i evidenciju u kompletan operativno-komercijalni sistem za otkupljivača / hladnjaču / organizatora proizvodnje.

## 1. Polazna ocena

AgriX već ima snažan transakcioni backbone:

- otkup i dokumentacioni lanac;
- kooperante, parcele i terenski deo;
- otpremnice, zbirne, prijemnice i fakture;
- sledljivost i palete;
- ambalažu kao ledger;
- novac, banku i SEF;
- izveštavanje, storno i operativne kontrole.

Zato sledeći veliki problem nije „još jedan ekran“ niti „još jedan izveštaj“.

Glavni funkcionalni rizik je da proizvod ostane veoma dubok u sredini poslovnog procesa, a da početak i kraj lanca ostanu van sistema.

Ciljni poslovni lanac treba da bude:

```
plan proizvodnje
    -> ugovor sa kooperantom
    -> očekivani otkup
    -> otkup
    -> kvalitet
    -> lot / operativni lager
    -> ugovor / porudžbina kupca
    -> rezervacija robe
    -> isporuka
    -> faktura
    -> banka / naplata
    -> profitabilnost
    -> management control
```

Danas je **sredina tog lanca najjača**. Najveće bele zone su oko ugovaranja, operativnog stock-a, prodajnog fulfillment-a, kvaliteta, planiranja i profitabilnosti.

---

## 2. Tier 1 — fundamentalne funkcionalne rupe

### 2.1 Ugovori i komercijala

Ovo je jedna od najvećih pravih praznina.

AgriX treba da zna ne samo **šta je urađeno**, nego i **šta je bilo dogovoreno**.

Capability treba da pokrije najmanje:

- ugovore sa kooperantima i kupcima;
- ugovorene količine;
- kulturu, sortu, klasu / kvalitet;
- cenovne formule ili ugovorenu cenu;
- rokove i periode važenja;
- avanse, bonuse, penale i druga komercijalna pravila;
- realizovanu i preostalu količinu;
- vezu ugovor -> otkup / isporuka / faktura.

Ključna poslovna pitanja:

- koliko je ugovoreno;
- koliko je realizovano;
- koliko još mora da se otkupi ili isporuči;
- koji ugovor je u riziku;
- koja roba je već komercijalno obavezana.

Bez ovog sloja AgriX dobro dokumentuje izvršenje, ali ne poseduje canonical sliku poslovne obaveze.

### 2.2 Operativni lager / stock management

**Sledljivost nije isto što i lager. Fiskalni lager nije isto što i operativni WMS/stock capability.**

Cilj je da sistem može pouzdano da odgovori:

> Koliko robe trenutno stvarno imam, gde se nalazi, kakvog je kvaliteta i koliko je slobodno za novu prodaju?

Minimalni capability:

- stock po objektu / lokaciji / komori;
- lot / batch identitet;
- ulaz, izlaz i interni premeštaj;
- inventura i korekcije;
- rezervisana, blokirana i slobodna količina;
- stanje po kulturi / sorti / klasi / kvalitetu;
- veza stock-a sa poreklom i dokumentima;
- veza sa preradom gde je primenljivo.

Primer:

```
Malina I klasa
Objekat A / Komora 3
Fizičko stanje:      37.8 t
Rezervisano:         12.0 t
Blokirano kvalitetom: 2.1 t
Slobodno:            23.7 t
```

Za hladnjače i ozbiljnije otkupljivače ovo je core capability, ne dodatak.

### 2.3 Prodaja / orders / fulfillment

Kupac ne treba da bude samo matični podatak ili posledica fakture.

Potrebna je prava prodajna strana procesa:

```
ponuda
 -> porudžbina
 -> rezervacija robe
 -> nalog za isporuku
 -> otpremnica
 -> faktura
 -> naplata
```

Capability treba da odgovori:

- šta je kupac naručio;
- kada mora biti isporučeno;
- šta je rezervisano;
- šta je stvarno isporučeno;
- šta kasni;
- šta još treba fakturisati;
- da li postoji dovoljno odgovarajuće robe za fulfillment.

Ovo zatvara drugu polovinu lanca koju otkup sam po sebi ne pokriva.

### 2.4 Quality / laboratorija / klasiranje

Sledeći nivo sledljivosti nije samo:

> od koga je roba?

nego:

> kakav je kvalitet te robe i mogu li to dokazati?

Capability može uključiti:

- parametre kvaliteta po kulturi;
- uzorke i rezultate analiza;
- Brix, vlagu, kalibar, nečistoće, temperaturu i druge relevantne mere;
- klasiranje i promenu klase;
- quarantine / release status;
- vezu rezultata sa lotom, kooperantom i prijemom;
- dokumente / slike / sertifikate;
- reklamacije povezane sa kvalitetom.

Ne treba unapred graditi laboratorijski LIMS, ali AgriX treba da ima jasan quality model ako cilja hladnjače, izvoz i ozbiljnije kupce.

---

## 3. Tier 2 — capability-ji koji značajno podižu vrednost proizvoda

### 3.1 Plan proizvodnje i očekivani otkup

AgriX treba da odgovori ne samo na pitanje šta se dogodilo, nego i:

> Šta očekujemo da će stići narednih dana i nedelja?

Po kooperantu / parceli / kulturi:

- površina;
- sorta;
- očekivani prinos;
- očekivani datum / period berbe;
- planirana količina;
- dnevni ili nedeljni plan prijema;
- realizacija plana;
- odstupanje.

Primer:

```
Kooperant A
Plan:               18,000 kg
Do danas:            9,400 kg
Forecast završetka: 15,800 kg
Odstupanje:          -12.2%
```

Ovaj sloj direktno hrani dispatch, kapacitete, hladnjaču, prodaju i cash planning.

### 3.2 Profitability / troškovna ekonomika

Ne treba praviti kompletno knjigovodstvo da bi AgriX odgovorio na pitanje koje vlasnika najviše zanima:

> Gde stvarno zarađujem, a gde samo pravim promet?

Potrebno je povezati:

- nabavnu cenu;
- transport;
- ambalažu;
- skladištenje;
- preradu;
- manipulaciju;
- bonuse / rabate;
- ostale direktno pripisive troškove;
- prodajnu cenu.

Iz toga treba izvesti landed cost i maržu po:

- kupcu;
- kulturi;
- sorti / klasi;
- lotu;
- otkupnom mestu;
- kooperantu gde ima smisla;
- periodu / kampanji.

### 3.3 AgriX Control Center

Control Center je najveći „next-level“ feature, ali **nije zamena za fundamentalne bele zone iz Tier 1**.

Njegov posao nije da pravi još podataka, nego da postojeće podatke pretvori u odluke.

Ne treba graditi dashboard sa mnogo grafikona. Treba graditi **exception-driven command center**.

Primer:

| Oblast | Status | Problem | Finansijski efekat | Akcija |
|---|---|---|---:|---|
| Otkup | crveno | 2 nezatvorene zbirne | 428k RSD | Otvori |
| Banka | narandžasto | 4 naloga čekaju | 1.28m RSD | Pregled |
| Kupci | crveno | 3 dospela računa | 912k RSD | Potraživanja |
| Roba | narandžasto | Manjak za isporuku | 2.3 t | Planiraj |
| Dokumenta | crveno | 5 nepotpunih tokova | — | Reši |
| Marža | narandžasto | Kupac X ispod minimuma | -84k RSD | Analiza |

Cilj:

> korisnik uđe u AgriX i za 30 sekundi zna šta danas zahteva pažnju, koliko je novca ili robe u riziku i gde treba da reaguje.

To pomera AgriX iz **system of record** ka **system of action**.

---

## 4. Tier 3 — enterprise maturity

Ovo su važni capability-ji, ali posle jezgra iz Tier 1 i Tier 2.

### Workflow / approvals / responsibilities

Model:

```
problem -> odgovorna osoba -> rok -> status -> rešeno
```

Primeri:

- odobriti nalog za plaćanje;
- rešiti nepotpun dokument;
- kontaktirati kupca zbog duga;
- pregledati quality exception;
- zatvoriti odstupanje inventure.

To pravi razliku između softvera koji zaposleni koriste i softvera **u kojem firma radi**.

### Claims / reklamacije / nesaglasnosti

- reklamacija dobavljaču ili od kupca;
- razlog i kategorija;
- lot / dokument / partner;
- količina pod sporom;
- finansijski efekat;
- slike / dokumenti;
- korektivna mera;
- status i ishod.

### Advanced planning

Kasnije:

- planiranje kapaciteta;
- dispatch;
- plan prijema;
- plan isporuka;
- konflikt kapaciteta;
- plan komora / objekata;
- scenario planning.

---

## 5. Šta ne treba automatski uvlačiti u AgriX

Ne treba od AgriX-a praviti generički ERP.

Sledeće oblasti imaju smisla samo ako realan customer workflow pokaže potrebu:

- pun HR;
- obračun zarada;
- general ledger / kompletno računovodstvo;
- generičko održavanje imovine;
- enterprise WMS funkcije bez potrebe tržišta;
- TMS nivo logistike;
- AI chatbot samo zato što je AI dostupan.

Princip:

> AgriX treba da poseduje capability koji je presudan za agricultural procurement / cold-storage / trading workflow, a da se na generičke poslovne sisteme integriše gde je to racionalnije.

---

## 6. Prioritet razvoja

Ako se gleda proizvod, a ne trenutni tehnički backlog:

### Tier 1 — zatvoriti poslovni lanac

1. **Ugovori i komercijala**
2. **Operativni lager / stock**
3. **Prodajne porudžbine / fulfillment**
4. **Quality**

### Tier 2 — povećati ekonomsku i management vrednost

5. **Plan / forecast otkupa**
6. **Profitability**
7. **Control Center**

### Tier 3 — povećati organizacionu zrelost

8. **Workflow / approvals**
9. **Claims / non-conformity**
10. **Advanced planning / dispatch / capacity**

Redosled unutar Tier 1 može zavisiti od prvih ciljnih kupaca. Za hladnjaču će lager i quality često biti ispred ugovora; za trgovca / organizatora proizvodnje ugovori i sales fulfillment mogu biti prvi.

---

## 7. Sales implikacija

Slabiji pitch:

> AgriX vodi otkup, kooperante, prijemnice, zbirne, fakture, SEF, sledljivost, bankovne naloge i izveštaje.

To zvuči kao feature checklist i kupac može da ga poredi sa kombinacijom Excel + knjigovođa + program za fakture.

Jači pitch, kada se zatvore bele zone:

> AgriX vodi ceo put robe i novca — od plana i ugovora sa kooperantom, preko otkupa, kvaliteta i lagera, do porudžbine kupca, isporuke, naplate i stvarne marže. U svakom trenutku pokazuje šta je ugovoreno, šta fizički postoji, šta je slobodno za prodaju, šta kasni i gde je novac u riziku.

To više nije samo digitalizacija administracije.

To je **kontrola poslovanja**.

---

## 8. Glavni zaključak

AgriX nema problem da je „premali“. Naprotiv — postojeći proizvod je funkcionalno dubok.

Najveći rizik je da dubina ostane koncentrisana na dokumentaciono-transakcioni deo otkupa, dok ključni komercijalni i management slojevi ostanu van sistema.

Zato sledeći veliki cilj nije broj novih funkcija nego **zatvaranje kompletnog poslovnog kruga**:

```
plan -> ugovor -> otkup -> kvalitet -> stock -> prodaja
     -> isporuka -> faktura -> naplata -> profit -> akcija
```

Kada taj krug postoji, Control Center i forecasting prestaju da budu „dashboard features“ i postaju prirodan komandni sloj nad celim poslovanjem.
