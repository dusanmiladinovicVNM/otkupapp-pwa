# Golden scenariji — specifikacija za pregled

> **Status: PREDLOG, nije implementirano.** Poslednja stavka PR2.
> Konstrukcija se pregleda **pre** implementacije — jer je ovde najlakše
> napraviti tačno onu grešku zbog koje scenariji i postoje.

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

### Zabranjeno: oblik implementacije

| Ne sme | Zašto |
|---|---|
| broj **redova** u tabeli | menja se u PR3 po definiciji |
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

Snapshot **ne sme** da čita tabele u sirovom obliku. Čita ih kroz iste
read-modele koje koristi aplikacija (`GetAmbalazeStanje`, `GetIsplataForOtkup`,
`ValidateZbirnaInvariant`, …), pa se u PR3 menja implementacija tih čitača, ne
sam scenario.

Primer izlaza — namerno bez ijednog imena tabele:

```
KOOPERANT KOOP-1
  predao          I=1000.00  II=200.00
  isplaceno       DA
  saldo           0.00
  ambalaza 12/1   -40
ZBIRNA 12/090926
  poslato         I=1000.00  II=200.00
  primljeno       I= 980.00  II=195.00
  kalo                 20.00       5.00
FAKTURA
  iznos           123456.00
  stavki          2
  nefakturisano   0
```

---

## 4) Predlog scenarija (18)

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

## 5) Tri stvari koje tražim da se potvrde pre implementacije

1. **Rečnik iz §2** — je li išta zabranjeno zapravo potrebno, ili išta
   dozvoljeno zapravo previše veže za današnju strukturu?
2. **B4 kao `KNOWN_FAIL`** — golden se piše na *ispravnu* vrednost i test pada
   dok PR6 ne popravi primary-row bug. Alternativa je pisati golden na *današnju
   pogrešnu* vrednost pa ga menjati u PR6 — što bi značilo da mreža kodifikuje
   bug. Predlažem prvo; treba potvrda jer uvodi namerno crven test.
3. **Obim** — 18 scenarija je gornja granica onoga što se održava. Ako je
   previše, prvo bih izostavio D2, G3 i C2.

---

## 6) Šta ovo NE pokriva

Izgled forme, štampu, PDF i ponašanje nad pravim podacima — to ostaje na
operateru (`.claude/rules/testovi.md` §7). Golden scenariji mere **poslovni
ishod**, ne prezentaciju.
