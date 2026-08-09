# Razvoj parsera za novu banku (Banka import)

Vodič za dodavanje podrške za **novu banku** u bankarski uvoz izvoda. Prati
postojeću multi-bank arhitekturu — **ne pravi paralelu**, ne diraj deljeni
integritet/staging.

> Povezano: `docs/production-runbook-banka-import-setup.md` (puštanje u rad),
> `docs/production-runbook-banka-novac.md` (mapiranje/incidenti).

---

## 1. Arhitektura (kako je uklopljeno)

```
PDF izvoda
  -> ExtractTextFromPdf          (pdftotext -raw -nopgbrk -enc UTF-8; DELJENO)
  -> ParseBankaIzvodForImport    (ORKESTRATOR; modBankaImport)
        DetectBank(lines)  -> "KOMERC" | "PROCREDIT" | "HALK" | "ALTA" | ...
        Select Case bankId -> rutira na parser TE banke (5 funkcija)
        4-nivo integrity check   (DELJENO — saldo vs parsirane sume/brojevi)
        17-kolonski staging       (DELJENO -> tblBankaImport)
  -> frmBankaImport               (auto-map po jakim kljucevima; DELJENO)
```

Za novu banku pišeš **samo jedan modul** (`modBanka<Ime>.bas`) sa 5 funkcija i
dodaješ **jednu granu** u dispatch. Sve nizvodno (integritet, staging, mapiranje)
je deljeno i **ne dira se**.

### Šta se NE dira (reuse > new)

| Deljeno | Gde |
|---|---|
| pdftotext runner | `modBankaImportParserPdfToText.ExtractTextFromPdf` / `ResolvePdfToTextExePath` |
| Orkestrator + integritet + staging | `modBankaImport.ParseBankaIzvodForImport` (4-nivo provera, 17 kolona) |
| Saldo tip (ugovor) | `Public Type BankIzvodSaldo` (u `modBankaImportParserPdfToText`) |
| Router | `modBankaImport.DetectBank` + `Select Case` |
| Test | `modBankaImport.Test_BankParse` (bank-agnostic) |

---

## 2. Ugovor parsera (5 funkcija po banci)

Svaka banka izlaže **istih 5 Public funkcija** (menjaš samo sufiks imena banke):

```vba
Public Function ExtractIzvodBroj<Ime>(ByRef lines() As String) As String
Public Function ExtractIzvodDatum<Ime>(ByRef lines() As String) As String
Public Function ExtractIzvodRacun<Ime>(ByRef lines() As String) As String
Public Function ExtractIzvodSaldo<Ime>(ByRef lines() As String) As BankIzvodSaldo
Public Function ParseBankaIzvod<Ime>(ByVal txt As String) As Variant
```

`ParseBankaIzvod<Ime>` vraća 2D niz `result(1..N, 1..10)` — **10 kolona po
transakciji**, tačno ovim redom (isto kao Komercijalna):

| Kol | Značenje |
|---|---|
| 1 | Datum izvoda |
| 2 | Datum transakcije (izvršenja) |
| 3 | Partner (naziv) |
| 4 | Račun partnera (`xxx-…-xx`) |
| 5 | **Zaduženje = isplata** |
| 6 | **Odobrenje = uplata** |
| 7 | Šifra plaćanja |
| 8 | Svrha |
| 9 | Poziv na broj (model) |
| 10 | Referenca (bankarska oznaka) |

`BankIzvodSaldo` (za integrity gate) mora dobiti: `PocetnoStanje`,
`UkupanDuguje`, `UkupanPotrazuje`, `ZavrsnoStanje`, `BrojNalogaZaduzenje`,
`BrojNalogaOdobrenje`, `parsed=True`.

Integritet (u orkestratoru) proverava, po izvodu:
1. `Pocetno + Uplate - Isplate == Novo stanje`
2. `suma Odobrenja == UkupanPotrazuje`
3. `suma Zaduzenja == UkupanDuguje`
4. `broj Odobrenja == BrojNalogaOdobrenje`, `broj Zaduzenja == BrojNalogaZaduzenje`

Ako neka pukne, uvoz batch-a se rollback-uje **glasno** — nema tihe korupcije.
Zato je integritet i **tvoj glavni test**: ako prođe, sume i brojevi se do na
paru slažu sa bančinim sopstvenim totalima.

---

## 3. Postupak korak-po-korak

### Korak 0 — sirovi tekst (uvek isti flagovi kao uvoz)
U aplikaciji: `Alt+F8 -> Diag_DumpFullPdfText` -> izaberi PDF nove banke ->
pored PDF-a nastane `<ime>.pdftext.txt` (isti izlaz koji parser vidi:
`pdftotext -raw -nopgbrk -enc UTF-8`). **Analiziraj taj tekst**, ne sam PDF.

### Korak 1 — pročitaj layout
Pre pisanja koda odgovori na:
- **Broj-format?** `1,234.56` (US, kao Komercijalna/ProCredit/Halk) ili `1.234,56`
  (srpski)? Diktira `ToNumber`/`IsAmount` — svaka banka svoj.
- **Zaglavlje:** gde su broj izvoda, datum, broj računa (koji marker).
- **STANJE/saldo blok:** koji labeli, koliko tokena u data-liniji, gde se nalazi
  (može biti i **u sredini** dokumenta, ne na kraju — vidi Halk).
- **Transakcija:** šta je početni marker (R.B.), koliko **datuma** po transakciji,
  gde su iznosi zaduženja/odobrenja, šifra, svrha, poziv, referenca.
- **Sekcije koje NE ulaze u saldo** (npr. Halk `NEIZVRŠENI NALOZI`) — moraju se
  **odseći** (bound na terminator tipa „Ukupno…"), inače integritet puca.

### Korak 2 — DetectBank fingerprint
U `modBankaImport.DetectBank` dodaj granu **iznad** `Case Else` default-a. Otisak
mora biti **ASCII, distinktivan** i pojavljivati se u svakom izvodu te banke, a NE
u drugim bankama (pazi da se ne poklopi sa nazivom banke koji se javlja kao
*partner* u tuđem izvodu):

```vba
If InStr(1, s, "<DISTINKTIVAN ASCII MARKER>", vbTextCompare) > 0 Then
    DetectBank = "<ID>"
    Exit Function
End If
```
Redosled: specifičniji fingerprinti prvi; `Case Else` = `"KOMERC"` (Komercijalna).

### Korak 3 — novi `modBanka<Ime>.bas`
Kopiraj `modBankaProCredit.bas` kao šablon (jednostavniji) ili `modBankaHalk.bas`
(kompleksniji: 2 datuma, sekcija za isključiti). Implementiraj 5 funkcija +
**lokalne, privatne** helpere (`ToNumber<Ime>`, `IsAmount<Ime>`, `IsDateLine<Ime>`,
`NormalizeSpaces<Ime>`, …). Helperi su privatni namerno — svaka banka ima svoj
format, pa se imena ne sudaraju i izmena jedne banke ne obara drugu.

### Korak 4 — grana u dispatch
U `ParseBankaIzvodForImport`, `Select Case bankId`, dodaj **iznad** `Case Else`:
```vba
Case "<ID>"
    brojIzvoda = ExtractIzvodBroj<Ime>(lines)
    datumIzvoda = ExtractIzvodDatum<Ime>(lines)
    brojRacuna = ExtractIzvodRacun<Ime>(lines)
    saldo = ExtractIzvodSaldo<Ime>(lines)
    txData = ParseBankaIzvod<Ime>(txt)
```

### Korak 5 — test (integritet = validacija)
`Ctrl+G` (Immediate) -> `Alt+F8 -> Test_BankParse` -> izaberi PDF. Očekuj:
```
=== DetectBank: <ID> ===
Izvod=…  Datum=…  Racun=…
Saldo: Pocetno=… Novo=… Duguje=… Potrazuje=…
--- OK: N transakcija ---
```
`--- OK: N ---` znači da je **integritet prošao**. `PARSE FAIL: … INTEGRITY FAIL /
PARSER MISMATCH` znači promašena/viška transakcija ili loš saldo — pogledaj koja
sekcija curi (npr. neisključena „na čekanju").

> **Kapija datuma (od v2.38.0 / RF-09):** pre saldo-provera, `ParseBankaIzvodForImport`
> traži da datum izvoda i datum **svake** transakcije prođu `TryParseDateValue`
> (round-trip: `30.02.` i sl. se odbijaju, ne prelivaju u sledeći mesec). Parser mora
> da vrati čist `dd.mm.yyyy` (ili `dd/mm/yyyy`) — spojene kolone, prazan datum ili
> `d.m.` bez godine daju `PARSE FAIL: … PARSER DATUM TRANSAKCIJE nije validan datum`
> uz redni broj transakcije.

### Korak 6 — kvalitet polja (integritet ovo NE pokriva)
Iz per-red dumpa proveri `racun`, `partner`, `svrha`, `poziv`, `referenca` —
njih integritet ne validira, a auto-map (`frmBankaImport`, Faza 7) koristi
**tekući račun** (primarno) i **poziv na broj** (sekundarno). Račun mora biti
čist; naziv banke (poreklo naloga) ne sme upasti u partnera/svrhu.

---

## 4. Naučene zamke (iz Komercijalna / ProCredit / Halk)

- **Broj-format nije prenosiv.** `ToNumber` koji radi za `1,234.56` ne radi za
  `1.234,56`. Prvo proveri format na uzorku.
- **Saldo blok varira.** Labeli prelomljeni preko linija; data-linija tipa
  „4 iznosa + 2 cela broja"; blok može biti **u sredini** izvoda (Halk, podnožje
  str. 1). Anchoruj na stabilan ASCII string (`"Prethodno stanje"`, `"duguje"`).
- **Sekcije koje ne ulaze u saldo.** Halk `NEIZVRŠENI NALOZI` ima **sopstvenu
  numeraciju** i istu strukturu — mora se odseći (bound na prvi „Ukupno na ra…").
  Bez toga integritet puca (višak transakcija).
- **Više datuma po transakciji.** Halk ima *datum izvršenja* + *datum prijema*.
  Ne pretpostavljaj jedan datum-pivot po transakciji.
- **Provizije/naknade.** Po-stavci naknade obično **ne** ulaze u dnevni promet
  (integritet to i potvrdi) — ne sabiraj ih u zaduženje/odobrenje.
- **Poziv na broj vs referenca.** Poziv (model, npr. `(00) …`, `003/26`) je za
  auto-map na otkup/fakturu; referenca (npr. Halk `0870011…`) je bankarska oznaka
  za dedup. Ne mešaj ih.
- **ASCII pravilo (obavezno).** VBA izvor mora ostati **100% ASCII**
  (vidi CLAUDE.md §4). Anchoruj na ASCII podstringove (`"Izvod za datum:"`,
  `"Ukupno na ra"`, `"duguje"`); ako baš moraš da poredis dijakritiku, koristi
  `ChrW(...)`. Nikad ne upisuj `š ž č ć đ` direktno u `.bas`. Posle izmene:
  `file modBanka<Ime>.bas` = „ASCII text", grep ne-ASCII = prazno.
- **Self-update.** Nov `modBanka<Ime>.bas` se **automatski kreira** kod klijenata
  kroz self-update (`VBComponents.Add`) — bez ručnog transfera. **Ne preimenuj**
  postojeće module (self-update ne briše stare -> „Ambiguous name"). Vidi
  `docs/SELF_UPDATE.md`.

---

## 5. Referentni primeri

| Banka | Modul | Karakteristike layout-a |
|---|---|---|
| Komercijalna | `modBankaImportParserPdfToText` | R.B. = broj `<=3` cifre; STANJE „Prethodno stanje"; blok-terminatori „Ukupno za račun" |
| ProCredit (`220-…`) | `modBankaProCredit` | R.B. bez tačke; datum-pivot; poziv `003/26`/`2026`; ref na kraju svrhe |
| Halkbank (`155-…`) | `modBankaHalk` | R.B. „N."; **2 datuma**; saldo u sredini; **NEIZVRŠENI sekcija se odseca**; ref `0870011…` |
| ALTA (`190-…`) | `modBankaAlta` | naslov **„IZVOD BR."** (fingerprint; Komercijalna/Halk = „Izvod broj"); **2 datuma** (knjiženja/prijema); STANJE „Prethodno stanje"; bound „PROMENE"…„Ukupno za ra"; **smer po „Obr. naknada"** (zaduženje = standalone iznos pre; odobrenje = `<iznos> <šifra>` posle); ref = 15-cifreni „Podaci za reklamaciju" |

Dispatch i test žive u `modBankaImport` (`DetectBank`, `Select Case`,
`Test_BankParse`).

---

## 6. Kontrolna lista (pre commita)

- [ ] `file src-vba/modBanka<Ime>.bas` = „ASCII text"; grep ne-ASCII = prazno.
- [ ] Balans `Function/End Function`, `Sub/End Sub` (statička provera).
- [ ] `DetectBank` fingerprint dodat iznad `Case Else`; `Case "<ID>"` u dispatch-u.
- [ ] `Test_BankParse` na uzorku: `--- OK: N transakcija ---` (integritet prošao).
- [ ] Per-red dump: `racun`/`partner`/`svrha`/`poziv`/`referenca` čisti.
- [ ] Commit na feature granu; test kroz `ImportAllVBA` -> `Compile`.
