# Golden scenariji — specifikacija za pregled

> **Status: grupa A implementirana (5 od 20), B–G čekaju.**
> `src-vba/modGoldenTests.bas`, suite `RunGoldenSuite`, goldeni u
> `tests/golden/`.
>
> Konstrukcija se pregleda **pre** implementacije — jer je ovde najlakše
> napraviti tačno onu grešku zbog koje scenariji i postoje. Pregled grupe A je
> uhvatio pet grešaka u samim scenarijima (§7).

---

## 1) Čemu služe

Sigurnosna mreža za PR3–PR12. Kad Zbirna postane header+stavke, moram da mogu da
dokažem da se **poslovni ishod nije promenio** — a ne da se tabela i dalje isto
zove.

Zato je jedini kriterijum kvaliteta scenarija:

> **Scenario koji bi morao da se menja u PR3 je napisan pogrešno.**

Ako test tvrdi „`tblZbirna` ima dva reda", on ne meri poslovanje nego zatečenu
strukturu — i moraću da ga prepišem baš u trenutku kad mi najviše treba
nepromenjen. To je ista greška kao test koji je nekad tvrdio da `GeneracijaID`
postoji: kodifikuje kompenzaciju umesto pravila.

---

## 1b) Izolacija ulaznog, ne samo izlaznog stanja

Rollback čisti ono što scenario **ostavi**. Ne čisti ono što je **zatekao**.

Prva verzija grupe A koristila je `KOOP-TEST-1` iz fixture-a i svih pet goldena
je javljalo `placeno 1000.00` — iako nijedan scenario ništa ne plaća. To je bio
zatečen avans, koji `SaveOtkupMulti_TX` automatski primeni. Izmena tog avansa u
fixture-u oborila bi goldene a da niko nije dirao poslovanje.

Zato scenariji koriste **sopstvene identitete bez transakcione istorije**:

```
KOOP-GLD-1   STA-GLD-1   VOZ-GLD-1   KUP-GLD-1
```

`GldPreduslov` to i **proverava** pre svakog scenarija: ako identitet već ima
otkup ili red u novcu, scenario staje sa imenovanom greškom. Golden test mora da
poseduje ceo poslovni ulaz koji utiče na izlaz, ne samo ID-eve koje je sam
napravio.

Gde je zatečeno stanje **predmet** testa (B2 — avans), scenario ga pravi sam, u
svojoj transakciji.

## 2) Rečnik tvrdnji — šta scenario SME da tvrdi

### Dozvoljeno: poslovne činjenice

| Kategorija | Primer tvrdnje |
|---|---|
| Roba | koliko kg klase I / II je kooperant predao; koliko je kupac primio; kolika je razlika (kalo) |
| Novac | saldo kooperanta; koliko je isplaćeno gotovinom; koliko avansom; da li je otkup **plaćen u celosti** |
| Ambalaža | saldo po entitetu i tipu, izveden iz ledgera |
| Status | koji **dokumenti** su aktivni, koji stornirani |
| Faktura | iznos; koje su stavke ušle; šta je ostalo nefakturisano |
| Sledljivost | iz kog otkupa potiče roba na konkretnoj fakturi |
| **Broj logičkih dokumenata** | „dva bloka su otišla na jednu otpremnicu" — jedan poziv writer-a = jedan dokument |

### Zabranjeno: oblik implementacije

| Ne sme | Zašto |
|---|---|
| broj **redova** u tabeli | menja se u PR3 po definiciji |
| broj **ID-eva** koje writer vrati | dvoklasni dokument danas vraća dva — to je broj redova prerušen u broj dokumenata |
| `GeneracijaID` bilo gde | mehanizam koji PR12 briše |
| poslovni broj kao **ključ pretrage** | A2 — broj je labela |
| indeks kolone, redosled kolona | to meri `VerifySchema`, ne poslovni test |
| format vraćenog ID-a (`"OTK-1 + OTK-2"`) | nestaje u PR5 |
| „primarni red" bilo koje vrste | primary-row hack je bug koji uklanjamo |

**Praktično pravilo:** ako se tvrdnja ne može izgovoriti kooperantu ili
knjigovođi, ne pripada golden scenariju.

---

## 3) Oblik

Reuse postojećeg: `AssertSnapshot(tekuci, imeGolden)` (`modTest`) piše/poredi
`tests/golden/<ime>.txt`. Mehanizam postoji i neiskorišćen je.

Svaki scenario je:

```
Seed        -> minimalni matični podaci + poslovni potez(i)
Snapshot    -> funkcija koja vraca POSLOVNE cinjenice kao tekst
AssertSnapshot -> poredjenje sa golden fajlom u gitu
```

### Golden Query Adapter

`GldSnapshot` i njegovi pomoćnici su **jedino mesto koje zna kako su podaci
složeni**:

```
Scenario  ->  Golden Query Adapter  ->  trenutni storage / read-model
```

Gde produkcioni read-model postoji, koristi se (`GetIsplataForOtkup`,
`SumOtpremniceByKlasa`, `IsZbirnaConsistent`, `GetAmbalazeStanje`). Za deo
činjenica ga **nema**, pa adapter čita tabelu direktno — to je svesna granica,
ne propust, i zato ovde piše umesto da se tvrdi da se koriste samo read-modeli.

**Scenario i golden fajl se nikad ne menjaju. Adapter sme.**

Stvarni izlaz (A2), namerno bez ijednog imena tabele i bez ijednog ID-a:

```
== A2 dvoklasni lanac ==
DOKUMENTI
  otkupa          1
  otpremnica      1
  zbirnih         1
  prijemnica      1
  faktura         1
OTKUP
  predao          I=1000.00  II=200.00
  vrednost        56000.00
  placeno         0.00
  isplaceno svi   NE
ZBIRNA
  poslato         I=1000.00  II=200.00
  primljeno       I=1000.00  II=200.00
  kalo            I=0.00  II=0.00
  invarijanta     OK
FAKTURA
  iznos           62000.00
```

`otkupa 1` iako dvoklasni otkup danas fizički daje dva reda — broji se
**poziv writer-a**, ne ID. Posle PR5 taj broj ostaje 1 i golden se ne menja.

---

## 4) Scenariji (20)

### A — Fresh Fruit Flow

| # | Scenario | Šta hvata |
|---|---|---|
| A1 | Jednoklasni otkup kroz ceo lanac do fakture | baseline; sve ostalo je odstupanje od njega |
| A2 | **Dvoklasni** otkup kroz ceo lanac | scenario koji PR5 najviše menja iznutra, a ishod mora ostati isti |
| A3 | Više blokova → jedna otpremnica | N:1 kardinalitet |
| A4 | Više otpremnica → jedna zbirna | agregat; §6.2 invarijanta |
| A5 | Kalo: poslato 1200, primljeno 1175 | razlika je poslovna činjenica, ne greška |

### B — Novac

| # | Scenario | Šta hvata |
|---|---|---|
| B1 | Gotovina pri otkupu | `tblNovac` → header |
| B2 | Avans primenjen na otkup | `ApplyAvansToOtkup` |
| B3 | **Delimična isplata** — `Isplaceno` ostaje prazno | prag „plaćeno u celosti" |
| B4 | Dvoklasni otkup plaćen u celosti | **danas ovo pada**: Klasa II nikad ne dobije `Isplaceno` (`modNovac.bas:1240` vs `modOtkup.bas:279`). Golden se piše na **ispravnu** vrednost, sa `KNOWN_FAIL` oznakom dok PR6 ne popravi. |

### C — Ambalaža

| # | Scenario | Šta hvata |
|---|---|---|
| C1 | OM izdaje prazne gajbe, kooperant vraća | ledger, obe noge |
| C2 | Saldo kroz ceo lanac | izvedeni saldo, ne kolona |

### D — Storno

| # | Scenario | Šta hvata |
|---|---|---|
| D1 | Storno otpremnice → zbirna se rekalkuliše | §6.2 posle mutacije |
| D2 | Storno prijemnice koja je fakturisana | kaskada |
| D3 | Storno dvoklasnog otkupa | **jedan** logički dokument, obe klase; novac i ambalaža poništeni |

### E — Ispravka

| # | Scenario | Šta hvata |
|---|---|---|
| E1 | Storno + reizdavanje pod **istim** poslovnim brojem | A9; danas traži `GeneracijaID`, posle PR12 ne sme |

### F — Faktura

| # | Scenario | Šta hvata |
|---|---|---|
| F1 | Faktura iz više prijemnica | 1:1 po stavci, puna količina |
| F2 | **Delimično fakturisanje** — Klasa I da, Klasa II ne | dokaz da je `Fakturisano` line-level (`DOCUMENT_HEADER_LINES.md` §4.4) |

### G — Ivični

| # | Scenario | Šta hvata |
|---|---|---|
| G1 | Samo Klasa II, bez Klase I | grana koju `hasKlasaI = False` menja |
| G2 | **Dva dokumenta sa istim poslovnim brojem** | G2 kapija ugovora; storno jednog ne dira drugi |
| G3 | Auto-hladnjača lanac | postojeća automatika ostaje funkcionalno identična |

---

## 5) Donete odluke

| Pitanje | Odluka |
|---|---|
| Rečnik tvrdnji (§2) | prihvaćen; dopunjen sa **brojem logičkih dokumenata** kao dozvoljenom činjenicom i **brojem ID-eva** kao zabranjenom |
| Obim | zadržava se **svih 20** |
| B4 | golden se piše na **ispravnu** vrednost — v. §5b |

### 5b) B4 ne sme da bude trajno crvena centralna kapija

`RunGoldenSuite` je `gate: True` i u podrazumevanom setu. Scenario koji trajno
pada pretvorio bi FULL u trajno crven — a provera koju operater nauči da
ignoriše ne štiti ništa. To je ista bolest kao placebo test, samo obrnuta.

Zato:

- golden za B4 se piše na **ispravnu** vrednost (kooperant plaćen u celosti →
  `isplaceno svi DA`) i **pregleda se**,
- ali se B4 **ne registruje** u `RunGoldenSuite` dok PR6 ne ukloni primary-row
  bug (`modNovac.bas:1240` računa po redu, `modOtkup.bas:279` piše novac samo na
  primarni red),
- PR6 ga registruje i time **dokazuje** da je bug popravljen.

Alternativa — pisati golden na današnju pogrešnu vrednost pa ga menjati u PR6 —
značila bi da sigurnosna mreža kodifikuje bug.

---

## 6) Šta ovo NE pokriva

Izgled forme, štampu, PDF i ponašanje nad pravim podacima — to ostaje na
operateru (`.claude/rules/testovi.md` §7). Golden scenariji mere **poslovni
ishod**, ne prezentaciju.

---

## 7) Šta je pregled grupe A uhvatio

Pet grešaka — sve u **scenarijima**, nijedna u sistemu. Zato se goldeni
pregledaju pre nego što se zaključaju.

| # | Greška | Posledica da je prošla |
|---|---|---|
| 1 | snapshot je čitao celu svesku po kooperantu | `predao 1720` iako scenario snima 1000; golden pada na svaku izmenu fixture-a |
| 2 | scenario je koristio `KOOP-TEST-1` iz fixture-a | `placeno 1000.00` u svih pet, iako nijedan ništa ne plaća — zatečen avans |
| 3 | format broja lokalno zavisan (`1720,00`) | golden pada na mašini sa drugom decimalnom oznakom; test meri Control Panel |
| 4 | ambalaža nije usklađena otpremnica ↔ zbirna | `invarijanta PUKLA` zabeležena kao da je sistem kriv |
| 5 | broj dokumenata brojan po ID-u | `otkupa 2` za jedan dvoklasni otkup — broj redova prerušen u broj dokumenata; golden bi se menjao u PR5 |

Greška 2 je najvažnija i vredi je pamtiti kao pravilo:

> **Rollback rešava ono što scenario ostavi iza sebe. Ne rešava ono što je
> zatekao.**
