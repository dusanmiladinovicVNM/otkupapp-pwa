# Odluka: šta se dešava sa novcem kad se otkup stornira ili ispravi

> **Status: OTVORENO — čeka operatera.** Ovo je poslovno pravilo, ne bug.
> Kod danas radi jednu stvar; dokument meri šta tačno, koje su alternative i
> šta svaka košta. Kad odluka padne, ovaj fajl postaje spec i test se okreće.
>
> Nađeno: 11.09.2026, u PR #308 (Otkup cutover, korak 5). Mereno, ne procenjeno.

---

## 1) Šta se dešava danas

```
StornoOtkup(otkupID)
   -> ResetNovacOtkupLink(otkupID)
         -> svakom tblNovac redu sa tim OtkupID-em:  OtkupID = ""
```

Novac se **ne stornira** — samo mu se skine veza za dokument. Namera je čitljiva
i razumna: roba je stornirana, ali novac je stvarno isplaćen, pa se ne sme
obrisati.

**Problem je šta se posle dešava sa tim odvezanim redom.**

### Tri tipa mogu biti vezana za otkup

| Tip | Ko ga veže | Šta znači |
|---|---|---|
| `VirmanFirmaKoop` | mapiranje izvoda (`modBankaMapiranje:3562`) | uplata firme kooperantu za konkretan blok |
| `VirmanAvansKoop` | `ApplyAvansToOtkup` (`modNovac:1645`) | avans naknadno primenjen na blok |
| `KesOtkupacKoop` | `SaveOMUlaz_TX` (`modDokumenta:7274`) | keš isplaćen kooperantu na otkupnom mestu |

### Ali samo JEDAN tip iko posle pokupi

Tri čitača, sva tri sa istim filterom `Tip = NOV_VIRMAN_AVANS_KOOP`:

```
modNovac:1624   ApplyAvansToOtkup             -- primena na nov dokument
modNovac:1982   GetKooperantUnallocatedAvans  -- "koliko avansa kooperant ima"
modNovac:2031   BuildKooperantUnallocatedAvansDict -- ekran isplata
```

Posledica, po tipu:

| Tip posle odvezivanja | Vidi ga dug? | Vidi ga avans? | Ishod |
|---|---|---|---|
| `VirmanAvansKoop` | ne (nema `OtkupID`) | **da** | vraća se u opticaj ✅ |
| `VirmanFirmaKoop` | ne | **ne** | **nevidljiv** ⚠ |
| `KesOtkupacKoop` | ne | **ne** | **nevidljiv** ⚠ |

Novac **nije izgubljen** — ostaje u kartici kooperanta (`modIzvestaj:2507`, ulazi
u saldo). Ali je ispao iz svake mašinerije koja odlučuje **šta se plaća**:
operater posle ispravke vidi nov dokument kao **pun dug**, a plaćeni iznos nigde
među avansima.

> **Nije uvedeno refaktorom.** Isto radi običan `StornoOtkup_TX` i radio je
> oduvek. Ispravka (`IspravkaOtkupa_TX`, PR #308) ga samo čini lakše dostižnim,
> jer sad postoji jedan potez koji stornira i odmah pravi naslednika.

---

## 2) Dva nalaza usput

### 2a) Storno JESTE povratan — ispravka nije

`modStorno:2176` već zna da je ovo osetljivo mesto:

```vb
' Zurnal: brisanje OtkupID->"" je jedini NEPOVRATNI deo storna otkupa.
' Zabelezi (NovacID, stari OtkupID -> "") PRE brisanja da undo re-linkuje.
JournalCell TBL_NOVAC, nvID, COL_NOV_OTKUP_ID, CStr(data(i, colOtkupID)), ""
```

Znači: **„Poništi storno" (`UndoStorno_TX`) vraća vezu.** Aparat postoji i radi.

Ono što ne postoji je **prenos na naslednika**: ispravka ne poništava storno nego
pravi nov dokument, pa žurnal nema kome da vrati vezu.

### 2b) `ResetNovacOtkupLink` postoji DVA puta, i samo jedan žurnalira

| Gde | Vidljivost | Žurnal |
|---|---|---|
| `modStorno:2156` | `Private` | **da** |
| `modNovac:1739` | `Public` | **ne** |

`StornoOtkup` (`modStorno:239`) zove **privatnu** — VBA prvo razrešava
modul-lokalno ime — pa produkcioni storno jeste žurnaliran. Javna verzija se
danas zove samo iz testa (`modNovacTests:457`), pa je opasnost **latentna, ne
živa**: prvi sledeći modul koji pozove javnu dobio bi tiho nepovratan storno.

`vba_check` ovo ne hvata: `DUPLIKAT` gleda dva **`Public`** imena, a ovde je jedno
`Private`. Nezavisno od odluke o novcu, ovo treba srediti.

---

## 3) Četiri puta, sa cenom

### A — Prenos na naslednika

`IspravkaOtkupa_TX` posle storna **preveže** oslobođene redove na nov `OtkupID`.

| | |
|---|---|
| **Dodiruje** | `modOtkup.IspravkaOtkupa_TX` (jedno mesto, unutar postojeće transakcije) |
| **Obim** | ~25 linija + test |
| **Za operatera** | ispravka je „isti posao, ispravljen papir" — novac ostaje gde je i bio |
| **Rizik** | prenosi se i novac koji se **ne odnosi** na novi iznos: ispravka sa 400 kg na 380 kg ostavlja preplatu, koju niko ne prijavljuje |
| **Ne rešava** | običan storno **bez** ispravke — tamo novac i dalje nestaje |

### B — Konverzija u avans

Storno oslobođenom redu **promeni tip** u `VirmanAvansKoop`, pa ga postojeća
mašinerija sama pokupi.

| | |
|---|---|
| **Dodiruje** | `ResetNovacOtkupLink` (obe kopije — v. §2b) |
| **Obim** | ~10 linija + test, ali **prethodno spajanje dve kopije** |
| **Za operatera** | novac se pojavi kao slobodan avans kooperanta; sledeći dokument ga automatski povuče |
| **Rizik** | **menja istorijsku činjenicu**: keš isplaćen na otkupnom mestu postaje „virmanski avans". Kartica kooperanta i izveštaji po kanalu plaćanja (`modIzvestaj:2503`) prikazuju pogrešan kanal |
| **Rešava** | i storno i ispravku, jednim pravilom |

### C — Fail-closed pre ispravke i storna

Otkup sa knjiženim isplatama se **ne može** stornirati ni ispraviti dok se novac
ne razveže ručno.

| | |
|---|---|
| **Dodiruje** | `StornoOtkup` + `IspravkaOtkupa_TX` |
| **Obim** | ~15 linija + testovi |
| **Za operatera** | **oduzima** operaciju koju danas ima; traži dva koraka umesto jednog |
| **Rizik** | najveći za svakodnevni rad — greška u unosu plaćenog otkupa postaje procedura |
| **Rešava** | ne rešava ništa — samo sprečava da se stanje napravi |

### D — Prošire se čitači

Tri čitača prestaju da filtriraju po tipu: **svaki nevezan red kooperanta** je
raspoloživ novac.

| | |
|---|---|
| **Dodiruje** | `modNovac:1624`, `:1982`, `:2031` |
| **Obim** | ~6 linija + testovi |
| **Za operatera** | odvezani novac se odmah vidi kao raspoloživ, bez obzira na kanal |
| **Rizik** | **najveći domenski**: `KesOtkupacKoop` bi postao „avans". Keš isplaćen na licu mesta nije avans nego zatvoren posao; kad bi ga `ApplyAvansToOtkup` povukao na tuđ dokument, novac bi „platio" nešto što nije |
| **Rešava** | i storno i ispravku, bez menjanja istorijske činjenice o kanalu |

---

## 4) Preporuka

**A + D, ali D suženo na tip — ne na „sve".**

Obrazloženje:

- **A** rešava pravi slučaj koji operatera boli — ispravka plaćenog otkupa — i
  radi na jednom mestu, unutar transakcije koja već postoji.
- **D suženo** znači: čitači primaju `VirmanAvansKoop` **i** `VirmanFirmaKoop`
  (oba su virmanski novac firme prema kooperantu, razlika je samo da li je u
  trenutku uplate bio poznat blok), a **`KesOtkupacKoop` ostaje isključen** jer
  keš na otkupnom mestu nije avans.
- **B** se odbacuje: laž o kanalu plaćanja je skuplja od problema koji rešava.
- **C** se odbacuje: oduzima operaciju umesto da popravi posledicu.

Uz to, **nezavisno od odluke**: spojiti dve kopije `ResetNovacOtkupLink` u jednu,
žurnaliranu (§2b).

Otvoreno pitanje koje A nosi i koje ti moraš zatvoriti: **kad ispravka smanji
iznos, preneseni novac postaje preplata.** Da li se ta preplata prijavljuje
operateru, ostavlja tiho, ili blokira ispravku?

---

## 5) Šta se dešava kad odluka padne

`Test_OTK_IspravkaNeGubiNovac` danas tvrdi **zatečeno** stanje:

```
Ispravka novac: nov dokument NE preuzima odvezanu isplatu
Ispravka novac: odvezana isplata NIJE slobodan avans
Ispravka novac: nov dokument je otvorena obaveza u punom iznosu
```

Te tri tvrdnje se **okreću**, ne brišu — test već stoji na pravom mestu i meri
pravu stvar. To je i razlog zbog kog je pisan u tom obliku umesto da se nalaz
ostavi u komentaru.
