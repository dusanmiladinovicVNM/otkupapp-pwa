# Claude Code — AGRIx VBA master: poslovna ambalažna kartica kooperanta

Ti si glavni developer za `dusanmiladinovicVNM/otkupapp-pwa`, koji vodi **master XLSM bez istorijskih podataka**. Uradi **proizvodnu, generičku** ispravku na najnovijem SHA/branchu, u skladu sa projektnim pravilima i testovima. Nemoj prepisivati privremene MIG QA fajlove niti menjati historijske podatke. Ne pretpostavljaj da je kod jednak starom SHA 66f7897 — proveri stvarni najnoviji kod.

## Problem i poslovni ugovor

AMB-10 `tblAmbalaza` s pravom knjiži `ULAZ_TUDJE_AMBALAZE`: `SpoljniSvet -> Kooperant` kao **tehničko pokriće deficita**, uz realni transfer `Kooperant -> Stanica` tipa `AMBALAZA_UZ_ROBU`. To knjiženje mora ostati u fizičkoj append-only knjizi i u računici fizičkog stanja i obaveze firme; ali **nije fizičko izdavanje praznih gajbica kooperantu**, ne sme da uveća **ulaz** niti poslovni saldo u finansijskoj/ambalažnoj kartici kooperanta.

Konkretni dokazni slučajevi iz arhivske 2026 (koristiti samo kao regresione scenarije, NE upisivati u master):
- `KOOP-00019`, 08.07.2026, `1/080726-16`: stvarni izlaz **23**, nema stvarnog ulaza **12**; taj +12 je automatski `ULAZ_TUDJE_AMBALAZE` i ne sme ući u poslovni promet.
- `KOOP-00019`, 11.07.2026, `OTK-01472`, `1/110726-11`: stvarni `AMBALAZA_UZ_ROBU` `Kooperant -> ST-00001`, **izlaz 3**; zasebni `ULAZ_TUDJE_AMBALAZE` `SpoljniSvet -> Kooperant`, **+3**, samo tehnički, NE poslovni ulaz. *U postojećem prikazu se vidi samo +3; proveri i zbog čega poslovni izlaz -3 nije prikazan.*
- Realne samostalne operacije izdavanja/povrata, reversi, storna, klase/tipovi ambalaže i međuperiodno početno stanje moraju ostati ispravno vidljivi.

## Obavezna arhitektura

1. **Ne menjaj** `tblAmbalaza`, upis, AMB-10 invarijante, istorijska knjiženja, `AmbSaldoNaloga`, `GetAmbalazeStanje`, `AmbSaldoPoNalogu` i obračun `AmbObavezaPartneru`. Fizički saldo i poslovni promet su dva različita pojma.
2. Napravi eksplicitan, dobro imenovan read model za **poslovna kretanja ambalaže kooperanta**. Klasifikuj po `VrstaKretanja` (i `StornoOd`/originalnoj vrsti za kontra-stav), ne po datumu, količini, `OtkupID` niti nazivu dokumenta. `ULAZ_TUDJE_AMBALAZE` je tehnički red, ne poslovni ulaz; i njegov storno mora biti uparen pravilno da se ne stvori fiktivan saldo. Kanonska čitanja svih kretanja ostaju neizmenjena ili kompatibilna.
3. `ReportKarticaAmbalaze` mora prikazati stvarne izdate/predate gajbice i njihove kontra-stavove, bez tehničkog pokrića. `ReportKarticaKooperanta` takođe mora koristiti **istu poslovnu definiciju**, izbegavajući dvostruko brojanje otkup-vezanih događaja. Proveri i sve druge prikaze KOOP poslovnog salda; ne menjaj obračun fizičkog salda u operativnim kapijama.
4. UI KPI `Saldo (kraj perioda)`, poslednji red, `UKUPNO`, straničenje, periodski početni saldo, PDF i prikaz moraju poticati iz iste poslovne projekcije. Nije prihvatljivo da KPI pokazuje 5, kartica 8 a stara evidencija -10 bez jasnog razloga.
5. **Fail-closed**: poslovni transfer bez potrebnog identiteta/vrednosti i dalje mora signalizirati anomaliju; ne sklanjaj celu grešku `On Error Resume Next`, ne izjednačavaj nepoznatu vrednost sa nulom.

## Testovi i acceptance

- Jedinični test koji unosi validna dva ledger događaja za OTK (`-3` realni + `+3` pokriće) i pokazuje: poslovna kartica daje `-3`; fizički ledger daje `0` neto; obaveza za tuđu ambalažu ostaje dostupna.
- Slučaj `-23` realni, `+12` tehnički: poslovna promena `-23`, ne `-11`.
- Jednostavan izdati/povratni revers, direktno izdavanje kooperantu, dva tipa ambalaže, storno poslovnog i storno tehničkog reda, međuperiodni saldo (događaji pre `datumOd`), bez dupliranja kroz finansijsku karticu.
- Precizno proveri zašto iz `tblAmbalaza` fizički red `OTK-01472` -3 nije vidljiv u postojećem `ReportKarticaAmbalaze` screenshotu. Filter tehničkog pokrića **nije dovoljan** ako ovaj -3 i dalje nedostaje.
- Potvrdi da nijedan `AMB-INV-*`, fizički saldo, generisanje otkupa/reversa, otpremnice, ili standardni izveštaji drugih entiteta nisu regresirali.
- Automatski testovi + VBA compile + ako je moguće Windows XLSM smoke. Dokaz, ne samo tvrdnja.

## Isporuka

Prikaži trenutno stanje repozitorijuma i SHA, precizan root cause, diff po fajlovima, nazive testova, rezultate, P1/P2 rizike i GO/NO-GO. Jedan fokusiran PR. Migrirane istorijske podatke ne diraj; sve projektuj tako da ispravno radi za potpuno novu sezonu 2027. Izbegavaj privremena arhivska `If 2026 Then` pravila u masteru.
