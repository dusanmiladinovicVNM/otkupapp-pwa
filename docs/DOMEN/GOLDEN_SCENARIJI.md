# Golden scenariji — specifikacija za pregled

> **Status: 13 scenarija implementirano i zaključano.**
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

## 4) Scenariji (13 zadržanih)

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
| B2 | **Pun avans** → `isplaceno svi DA` | `ApplyAvansToOtkup` zatvara otkup |
| B3 | **Delimičan avans** → `isplaceno svi NE` | prag „plaćeno u celosti" |

> B1 i B4 su uklonjeni: merili su keš uz otkupni list, koji u domenu ne postoji
> (§8).

> **Izbačeni odlukom:** C1, C2 (ambalaža — pokriva `RunStornoTestSuite` i
> sekcija `AMBALAZA` u svakom golden-u), E1 (ispravka), F1 (faktura iz više
> prijemnica — pokriva `RunBusinessFlowProSuite`), G3 (auto-hladnjača).

### D — Storno

| # | Scenario | Šta hvata |
|---|---|---|
| D1 | Storno otpremnice → zbirna se rekalkuliše | §6.2 posle mutacije |
| D2 | Storno prijemnice koja je fakturisana | kaskada |
| D3 | Storno dvoklasnog otkupa | **jedan** logički dokument, obe klase; novac i ambalaža poništeni |

### F — Faktura

| # | Scenario | Šta hvata |
|---|---|---|
| F2 | **Delimično fakturisanje** — Klasa I da, Klasa II ne | dokaz da je `Fakturisano` line-level (`DOCUMENT_HEADER_LINES.md` §4.4) |

### G — Ivični

| # | Scenario | Šta hvata |
|---|---|---|
| G1 | Samo Klasa II, bez Klase I | grana koju `hasKlasaI = False` menja |
| G2 | **Dva dokumenta sa istim poslovnim brojem** | G2 kapija ugovora; storno jednog ne dira drugi |

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
| 6 | `DOKUMENTI` je brojao **pozive testa** (`m_nOtp = m_nOtp + 1`) | tautologija: A3 je dokazivao „test je jednom pozvao `GldOtpremnica`", ne „sistem je napravio jednu otpremnicu"; bug koji od jednog poziva napravi dve otpremnice po 500 kg ostavio bi agregat isti i test **zelen** |
| 8 | seed rezervisanih identiteta bio fail-open | zatečen `STA-GLD-1` sa drugim nazivom ili `Aktivan=Ne` prihvatio bi se — scenario opet ne poseduje ceo svoj ulaz |
| 7 | preduslov je proveravao samo identitet, ne ključ scenarija | zatečena otpremnica sa `BrojZbirne = GLD-A1` ušla bi u rezultat i kad kooperant nema nijedan stari otkup |

Dve greške vrede da se pamte kao pravila:

> **Rollback rešava ono što scenario ostavi iza sebe. Ne rešava ono što je
> zatekao.** (greška 2, dopunjeno greškom 7)

> **Oracle mora da meri šta je sistem napravio, ne šta je test nameravao.**
> (greška 6)

### Dokaz da oracle zaista grize

`DOKUMENTI` sada čita sistem: distinct `GeneracijaID`, a gde ga red ne nosi
(`tblOtkup`) — distinct poslovni broj, sve u opsegu scenarija.

Sabotaža: A3 pravi **dve** otpremnice po 500 kg umesto jedne od 1000.

```
poslato 1000.00          <- agregat NEPROMENJEN
golden [otpremnica 1] vs tekuci [otpremnica 2]   <- pada OVDE
```

Posle PR5 adapter postaje `COUNT(DISTINCT <Doc>ID)`; golden ostaje isti.

### Ugovor izolacije, u celini

```
PRE      GLD identiteti          ne postoje
         transakciona istorija   ne postoji
         ključ scenarija         ne postoji
SCENARIO sam napravi sve što mu treba
POSLE    rollback -> isto stanje kao PRE
```

Sva tri uslova su **fail-fast** i imenovana. Dokazano sabotažom: dupli `GldSeed`
u istoj transakciji daje

```
rezervisani identitet STA-GLD-1 vec postoji u tblStanice (1)
  -- scenario ne poseduje svoj ulaz
```

Idempotentnost se namerno **ne** traži: svaki scenario radi u svojoj transakciji
i rollback-uje se, pa identiteti na početku uvek ne postoje. Ako postoje, to je
nalaz — ili je prethodni rollback zakazao, ili ime više nije rezervisano.

---

## 8) Keš se ne vezuje za otkupni list — B1 i B4 uklonjeni

**Poslovno pravilo:** keš nikada ne ide uz otkupni list. Ekran otkupnog lista u
novom UI-ju nema polje za novac, a `modOtkupUnos` inicijalizuje `p("novac") = 0#`
i **nigde ga ne menja** — jedine druge reference na taj ključ su u
`modNovacUnos`, što je zaseban rečnik novčanog ekrana.

Posledice:

| Šta | Ishod |
|---|---|
| `SaveOtkupMulti_TX(novac)` parametar | mrtav iz UI putanje |
| `SaveNovac(otkupID:=primaryID)` u `modOtkup.bas:279` | **mrtav kod**, ne bug koji treba popraviti |
| kolone `Novac`, `PrimalacNovca` na `tblOtkup` | uvek 0 / prazno → brišu se u refaktoru |
| B1 (keš pri otkupu) | **uklonjen** — merio je putanju koja u domenu ne postoji |
| B4 (dvoklasni plaćen kešom) | **uklonjen** iz istog razloga |
| B3 | prespecificiran na **delimičan avans** |

Čitaoci kolone `Novac` koji ostaju: `modOtkup.GetSaldoByStation`,
`modOtkupBlok:945` (štampa bloka), `modDokumenta:3696` — svi vide nulu i idu u
čišćenje zajedno sa kolonom.

### Ispravka ranije tvrdnje

Prva verzija ovog odeljka je tvrdila da `UpdateOtkupStatus` nije pozvan sa
putanje otkupa uopšte. **Netačno** — mereno `awk` opsegom koji je sekao pre
kraja procedure. `ApplyAvansToOtkup` ga zove na `modNovac.bas:1642`
(`If preostalo <= 0 Then UpdateOtkupStatus otkupID`), unutar sebe.

Tačno je uže: `Isplaceno` se postavlja kroz **avans**, banku i novčani modul —
ne kroz `novac` parametar otkupa. B2 to i dokazuje: pun avans 50 000 daje
`isplaceno svi DA`, delimičan 20 000 u B3 daje `NE`.

---

## 9) Šta svaki zaključan golden dokazuje

| Golden | Ključna tvrdnja |
|---|---|
| A1 | pun lanac do fakture: 1 dokument svake vrste, faktura 55000 |
| A2 | dvoklasni: **1** logički otkup (ne 2), ishod isti kao jednoklasni |
| A3 | 2 otkupa → **1** otpremnica; sabotaža sa 2×500 kg pada na kardinalnosti |
| A4 | **2** otpremnice → 1 zbirna, invarijanta OK |
| A5 | kalo 25 kg je poslovna činjenica, ne greška |
| B2 | pun avans 50 000 → `avansom 50000`, `isplaceno svi DA` |
| B3 | delimičan avans 20 000 → `isplaceno svi NE` (prag nije dostignut) |
| D1 | storno otpremnice: `aktivnih 0 / storniranih 1`, zbirna ostaje bez izvora |
| D2 | storno fakturisane prijemnice — kaskada |
| D3 | storno dvoklasnog otkupa gasi **jedan** logički dokument, obe klase |
| F2 | fakturisano `I=DA / II=NE` — dokaz da je `Fakturisano` line-level |
| G1 | samo Klasa II prolazi ceo lanac |
| G2 | isti broj, dva dokumenta: storno jednog ne dira drugi (`aktivnih 1 / storniranih 1`) |
