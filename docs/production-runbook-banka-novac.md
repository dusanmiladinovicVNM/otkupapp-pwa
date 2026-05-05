# Production runbook: BankaImport, BankaMapiranje i Novac

Status: **operativni runbook za incidente oko bankarskih izvoda, mapiranja uplata/isplata, avansa i finansijskog stanja.**

Aplikacija: **OtkupApp / AgriX**
Domen: **PDF bankarski izvod → `tblBankaImport` staging → `tblNovac` finansijska knjiga → `tblFakture` / `tblOtkup` statusi**
Glavni kod: `src-vba/modBankaImport.bas`, `src-vba/modBankaMapiranje.bas`, `src-vba/modNovac.bas`, `src-vba/frmBankaImport.frm`

---

## 1. Kada korisnik kaže problem

Tipični incidenti:

* “Uvezao sam izvod, ali uplata nije legla na fakturu.”
* “Uplata kupca je otišla u avans, a trebalo je na fakturu.”
* “Isplata kooperantu je vezana za pogrešan otkup.”
* “Banka stavka je pogrešno mapirana na partnera.”
* “PDF izvoda je nestao iz inbox-a.”
* “Banka import kaže da nema otvorenih stavki, ali uplata nije knjižena.”
* “Isti izvod je uvezen dva puta.”
* “PartnerMap je naučio pogrešnog partnera.”
* “Faktura je plaćena, ali status i dalje nije plaćeno.”
* “Kooperant je plaćen, ali otkup stoji neisplaćen.”

Prvo pravilo:

> Ne mapiraj ponovo i ne pravi ručni `NOV-*` red dok ne utvrdiš `BIM-*` red, njegov `Obradjeno` status i da li već postoji povezan `NOV-*` red.

Minimalni podaci koje operator mora da prikupi:

```text
BankaImportID / BIM-ID:
BrojDokumenta / broj izvoda:
IzvorFajl:
DatumIzvoda:
DatumTransakcije:
Partner:
PartnerKonto:
Uplata:
Isplata:
PozivNaBroj:
BankaReferenz:
Obradjeno:
Stornirano:
Ako je već mapirano: NovacID / NOV-ID:
Ako je kupac: KupacID, FakturaID:
Ako je kooperant: KooperantID, OtkupID / BrojBloka:
Ako je OM: OMID / StanicaID:
```

---

## 2. Source of truth: gde se gleda

### 2.1. Prvo mesto: `frmBankaImport`

Otvoriti formu **Banka import**.

Forma prikazuje otvorene stavke iz `GetBankaImportOpen()` i pokazuje:

* `BIM ID`;
* datum transakcije;
* partnera;
* poziv na broj;
* uplatu;
* isplatu;
* status `Obradjeno`;
* detalje: opis, svrha, poziv na broj;
* auto-preview: kome bi sistem mapirao red i koji tip knjiženja bi koristio.

Dostupne UI akcije:

| Dugme         | Kada se koristi                                        | Šta radi                                                                                 |
| ------------- | ------------------------------------------------------ | ---------------------------------------------------------------------------------------- |
| Auto jedan    | kada operator želi auto-map samo jedne stavke          | `AutoMapBankaImportRow_TX(bimID)`                                                        |
| Auto sve      | kada se mapiraju sve otvorene stavke                   | `AutoMapAllBankaImport_TX()`                                                             |
| Sačuvaj ručno | kada auto-preview nije dovoljan ili treba ručna odluka | `MapBankaImportAsKupac_TX`, `MapBankaImportAsKooperantBlock_TX`, `MapBankaImportAsOM_TX` |
| Skip          | kada stavka ne treba u `tblNovac`                      | `SkipBankaImportRow_TX(bimID)`                                                           |
| Osveži        | reload otvorenih stavki                                | `LoadBankaRows`                                                                          |

### 2.2. Drugo mesto: `tblBankaImport`

Ovo je staging tabela za stavke iz bankarskog PDF izvoda.

Obavezno proveriti:

| Kolona             | Značenje                                   |
| ------------------ | ------------------------------------------ |
| `BankaImportID`    | interni ID staging stavke, npr. `BIM-*`    |
| `BrojDokumenta`    | broj izvoda                                |
| `DatumIzvoda`      | datum izvoda                               |
| `BrojRacuna`       | račun izvoda                               |
| `DatumTransakcije` | datum bankarske transakcije                |
| `Partner`          | naziv partnera iz banke                    |
| `PartnerKonto`     | račun partnera                             |
| `Opis`             | opis transakcije                           |
| `Uplata`           | priliv novca                               |
| `Isplata`          | odliv novca                                |
| `Valuta`           | obično RSD                                 |
| `PozivNaBroj`      | poziv na broj, često ključ za fakturu/blok |
| `SvrhaPlacanja`    | svrha plaćanja                             |
| `BankaReferenz`    | bankarska referenca, važna za dedupe       |
| `IzvorFajl`        | PDF fajl iz koga je red došao              |
| `ImportVreme`      | vreme import-a                             |
| `Obradjeno`        | ``, `Da`, `Skip`, `Error`                  |
| `Stornirano`       | ako je staging red storniran/isključen     |

### 2.3. Treće mesto: `tblNovac`

Ovo je finansijska knjiga. Svaka mapirana bankarska stavka mora imati `NOV-*` red ili mora biti jasno označena kao `Skip` / `Error`.

Obavezno proveriti:

| Kolona          | Značenje                                                   |
| --------------- | ---------------------------------------------------------- |
| `NovacID`       | interni finansijski ID, npr. `NOV-*`                       |
| `BrojDokumenta` | broj izvoda / dokumenta                                    |
| `Datum`         | datum knjiženja                                            |
| `Partner`       | naziv partnera za finansijsku knjigu                       |
| `PartnerID`     | `KupacID`, `StanicaID` ili drugi partner ID                |
| `EntitetTip`    | `Kupac`, `OM`, eventualno drugi tip                        |
| `OMID`          | stanica/otkupno mesto ako je relevantno                    |
| `KooperantID`   | kooperant za isplate                                       |
| `FakturaID`     | faktura kojoj je uplata vezana                             |
| `VrstaVoca`     | za segmentaciju uplata/isplata                             |
| `Tip`           | finansijski tip: uplata, avans, virman, itd.               |
| `Uplata`        | priliv                                                     |
| `Isplata`       | odliv                                                      |
| `Napomena`      | mora sadržati BIM trag: `BIM:<id>; Ref:<...>; Konto:<...>` |
| `Stornirano`    | isključenje finansijskog reda                              |
| `OtkupID`       | otkup kome je isplata vezana                               |
| `OsirocenoOD`   | indikator orphan stanja ako postoji                        |

### 2.4. Četvrto mesto: `tblPartnerMap`

`tblPartnerMap` povezuje bankarski naziv partnera sa internim partnerom.

Obavezno proveriti:

| Kolona       | Značenje                   |
| ------------ | -------------------------- |
| `BankaName`  | naziv iz bankarskog izvoda |
| `PartnerID`  | interni partner ID         |
| `EntitetTip` | `Kupac`, `Kooperant`, `OM` |
| `OMID`       | stanica ako je relevantno  |

Ovo je rizično jer pogrešna mapa pravi ponavljajuće pogrešno auto-mapiranje.

### 2.5. Fajl sistem

Banka PDF folderi:

```text
APP_BANKA_INBOX      = novi PDF izvodi za import
APP_BANKA_PROCESSED  = uspešno parsirani PDF-ovi
APP_BANKA_ERROR      = PDF-ovi koje sistem nije mogao da parsira/importuje
```

Kod pomera PDF:

* prazan tekst ili parse failure → `APP_BANKA_ERROR`;
* uspešan parse/save → `APP_BANKA_PROCESSED`;
* greška tokom import-a → pokušava pomeranje u `APP_BANKA_ERROR`.

---

## 3. Koji ID pratiš

Uvek zajedno prati:

1. `BankaImportID` / `BIM-*` — staging transakcija iz izvoda.
2. `IzvorFajl` — PDF izvod.
3. `BrojDokumenta` — broj bankarskog izvoda.
4. `BankaReferenz` — bankarska referenca; najjači dedupe signal kada postoji.
5. `PozivNaBroj` — često poslovni ključ za fakturu/blok.
6. `NovacID` / `NOV-*` — finansijski red koji je nastao mapiranjem.
7. `FakturaID` — ako je uplata vezana na kupca/fakturu.
8. `OtkupID` — ako je isplata vezana na kooperantski otkup.
9. `PartnerMap` par: `BankaName → PartnerID / EntitetTip / OMID`.

Incident ticket minimum:

```text
BIM-ID:
NOV-ID:
IzvorFajl:
BrojDokumenta:
BankaReferenz:
PozivNaBroj:
Partner iz banke:
Uplata:
Isplata:
Obradjeno:
Mapirano kao: Kupac / Kooperant / OM / Skip / Error
PartnerID:
FakturaID:
OtkupID:
Tip knjiženja:
Operator:
Odluka:
```

---

## 4. Normalan tok: PDF izvod → `tblBankaImport`

Normalan import tok:

1. Operator stavi PDF izvod u `APP_BANKA_INBOX`.
2. Pokrene `ImportBankaInbox_TX()`.
3. Sistem pravi TX snapshot `tblBankaImport`.
4. Sistem prolazi sve `*.pdf` u inbox-u.
5. Za svaki PDF:

   * čita tekst iz PDF-a;
   * izvlači broj izvoda, datum izvoda i broj računa;
   * parsira transakcije;
   * deduplikuje redove;
   * upisuje nove redove u `tblBankaImport`;
   * pomera PDF u `APP_BANKA_PROCESSED`.
6. Ako PDF nema tekst ili parse ne uspe, ide u `APP_BANKA_ERROR`.
7. Ako dođe do greške, TX rollback vraća `tblBankaImport` snapshot.

Važno:

> PDF pomeranje je file-system side effect. Excel TX rollback vraća tabelu, ali ne mora automatski vratiti PDF u inbox. Zato kod incidenta uvek proveri i tabelu i folder.

---

## 5. Normalan tok: `tblBankaImport` → `tblNovac`

### 5.1. Dolazna uplata kupca

Uslov:

```text
Uplata > 0
Isplata = 0
```

Mogući ishod:

* ako se jednoznačno nađe kupac i faktura → `NOV_KUPCI_UPLATA`, `FakturaID` popunjen;
* ako se nađe kupac, ali ne i faktura → `NOV_KUPCI_AVANS`, `FakturaID` prazan;
* ako se ne nađe kupac → `BIM.Obradjeno = Error`.

Efekat:

* pravi `NOV-*` red;
* `BIM.Obradjeno = Da`;
* eventualno dodaje/koristi `tblPartnerMap`;
* ako postoji `FakturaID`, poziva `UpdateFakturaStatus`.

### 5.2. Odlazna isplata kooperantu

Uslov:

```text
Isplata > 0
Uplata = 0
```

Mogući ishod:

* ako se nađe kooperant i otvoreni otkup/blok → `NOV_VIRMAN_FIRMA_KOOP`, `OtkupID` popunjen;
* ako se nađe kooperant, ali nema otvorenih otkupa → `NOV_VIRMAN_AVANS_KOOP`, `OtkupID` prazan;
* ako se koristi blok, sistem raspodeljuje iznos po kandidatima i eventualni višak knjiži kao avans;
* ako se ne nađe kooperant → `BIM.Obradjeno = Error`.

Efekat:

* pravi jedan ili više `NOV-*` redova;
* povezuje isplate na `OtkupID` kada je moguće;
* poziva `UpdateOtkupStatus`;
* `BIM.Obradjeno = Da` ako je mapiranje uspelo.

### 5.3. OM / stanica

Mogući ishod:

* knjiženje prema `OMID` / `StanicaID`;
* tip obično `NOV_KES_FIRMA_OTKUPAC` ili drugi OM tok;
* `BIM.Obradjeno = Da`.

### 5.4. Skip

`Skip` se koristi kada bankarska stavka ne treba da postane `tblNovac` red.

Primeri:

* interna bankarska naknada koja se ne vodi u ovom modulu;
* test/dupli red koji je već ručno rešen;
* stavka van poslovnog scope-a aplikacije.

Skip mora imati poslovno/operatersko objašnjenje u incident ticket-u.

---

## 6. Statusi `tblBankaImport.Obradjeno`

| Status  | Značenje                           | Operator sme                 |
| ------- | ---------------------------------- | ---------------------------- |
| prazno  | otvoreno, nije obrađeno            | auto/ručno mapirati ili skip |
| `Da`    | mapirano u `tblNovac` ili obrađeno | ne mapirati ponovo           |
| `Skip`  | svesno preskočeno                  | ne mapirati bez odluke       |
| `Error` | mapiranje nije uspelo              | analizirati i ručno rešiti   |

Pravilo:

> `Obradjeno = Da` je hard stop za ponovno mapiranje dok se ne dokaže da povezani `NOV-*` red ne postoji ili je storniran uz odobrenje.

---

## 7. Dedupe pravila za import izvoda

Sistem smatra bankarsku stavku duplikatom ako:

1. isti `BrojDokumenta` već postoji; i
2. ako `BankaReferenz` postoji, ona se poklapa; ili
3. ako `BankaReferenz` ne postoji, poklapaju se:

   * `DatumTransakcije`;
   * `Uplata`;
   * `Isplata`;
   * `Partner`.

Operativno značenje:

* isti PDF može biti ponovo ubačen, ali postojeći redovi ne bi trebalo da se dupliraju;
* ako banka menja format/reference, dedupe može omanuti;
* ako isti partner ima dve identične transakcije istog dana bez reference, postoji rizik false duplicate-a.

Runbook pravilo:

> Ako korisnik tvrdi da fali jedna od dve identične uplate, proveri `BankaReferenz`. Ako reference nema, ručno uporedi ceo izvod pre zaključka da je duplikat.

---

## 8. Standardni incident flow

### Korak 1: Nađi `BIM-*` red

U `frmBankaImport` ili `tblBankaImport` pronađi red po:

```text
BIM-ID
IzvorFajl
BrojDokumenta
BankaReferenz
DatumTransakcije + Partner + iznos
PozivNaBroj
```

Zapiši:

```text
BIM-ID:
Obradjeno:
Stornirano:
IzvorFajl:
BankaReferenz:
Partner:
Uplata:
Isplata:
PozivNaBroj:
SvrhaPlacanja:
```

### Korak 2: Proveri da li već postoji `NOV-*`

Traži u `tblNovac` po:

* `Napomena` sadrži `BIM:<BIM-ID>`;
* `BrojDokumenta`;
* `Datum`;
* `Partner`;
* iznos `Uplata`/`Isplata`;
* `FakturaID` ili `OtkupID`.

Zapiši:

```text
NOV-ID:
Tip:
PartnerID:
EntitetTip:
FakturaID:
OtkupID:
Uplata:
Isplata:
Stornirano:
Napomena:
```

### Korak 3: Klasifikuj problem

| Signal                                    | Kategorija              | Sledeći korak                                                   |
| ----------------------------------------- | ----------------------- | --------------------------------------------------------------- |
| PDF je u `APP_BANKA_ERROR`                | parse/import problem    | proveri tekst PDF-a i parse grešku                              |
| PDF je u `PROCESSED`, nema `BIM-*` redova | TX/file inconsistency   | proveri rollback/log/journal; vrati PDF u inbox ako je bezbedno |
| `BIM.Obradjeno = prazno`                  | čeka mapiranje          | koristi preview, zatim auto/ručno mapiranje                     |
| `BIM.Obradjeno = Error`                   | auto-map failed         | ručno mapiranje ili master data fix                             |
| `BIM.Obradjeno = Skip`                    | svesno preskočeno       | samo uz poslovnu odluku menjati                                 |
| `BIM.Obradjeno = Da`, nema `NOV-*`        | nekonzistentno          | proveri rollback/journal; tehnički owner                        |
| `NOV-*` postoji, ali pogrešan partner     | pogrešno mapiranje      | ne praviti novi red; stornirati/korekcija uz odluku             |
| `NOV-*` postoji, ali pogrešan `FakturaID` | pogrešno vezana uplata  | finansijska korekcija, UpdateFakturaStatus                      |
| `NOV-*` postoji, ali pogrešan `OtkupID`   | pogrešno vezana isplata | finansijska korekcija, UpdateOtkupStatus                        |
| `PartnerMap` pogrešan                     | sistemski rizik         | ispraviti mapu pre daljeg auto-mapiranja                        |

### Korak 4: Izaberi dozvoljenu akciju

| Stanje                                    | Dozvoljena akcija                     |
| ----------------------------------------- | ------------------------------------- |
| `BIM.Obradjeno = prazno`, preview tačan   | Auto jedan / Auto sve                 |
| `BIM.Obradjeno = prazno`, preview netačan | ručno mapiranje                       |
| `BIM.Obradjeno = Error`                   | ručno mapiranje posle ispravke uzroka |
| `BIM.Obradjeno = Skip`                    | ne dirati bez odluke                  |
| `BIM.Obradjeno = Da`, `NOV-*` tačan       | zatvoriti incident                    |
| `BIM.Obradjeno = Da`, `NOV-*` pogrešan    | korekcija/storno, ne remap naslepo    |
| `BIM.Obradjeno = Da`, `NOV-*` ne postoji  | tehnički recovery                     |

---

## 9. Retry i remap pravila

### 9.1. Kada sme mapiranje

Mapiranje je dozvoljeno ako:

* `BIM.Obradjeno` je prazno; ili
* `BIM.Obradjeno = Error`, a uzrok je ispravljen; i
* ne postoji validan `NOV-*` red za isti `BIM-*`; i
* operator je proverio preview.

### 9.2. Kada ne sme mapiranje

Ne mapirati ako:

* `BIM.Obradjeno = Da`;
* `BIM.Obradjeno = Skip` bez odluke;
* postoji `NOV-*` red sa istim `BIM:<id>` u napomeni;
* postoji `NOV-*` red istog iznosa/datuma/partnera koji je verovatno nastao iz istog izvoda;
* problem je pogrešan `PartnerMap`, dok mapa nije popravljena;
* postoji poslovni spor oko toga da li je uplata avans ili plaćanje fakture.

### 9.3. Kako se radi korekcija pogrešnog mapiranja

Ne raditi “još jedan NOV red” da bi se stanje poravnalo, osim ako finansijski owner to eksplicitno odobri.

Standardno:

1. Identifikuj pogrešni `NOV-*` red.
2. Identifikuj `BIM-*` red.
3. Proveri da li je `NOV-*` već uticao na `FakturaID` ili `OtkupID` status.
4. Ako je potrebno, stornirati pogrešni `NOV-*` red ili ga ručno korigovati po definisanoj proceduri.
5. Pokrenuti odgovarajući status update:

   * `UpdateFakturaStatus(fakturaID)`;
   * `UpdateOtkupStatus(otkupID)`.
6. Tek zatim mapirati ispravno ili kreirati korekcioni finansijski red.
7. Popraviti `PartnerMap` ako je ona uzrok.

Ako u aplikaciji još nema eksplicitne “Undo bank mapping” procedure, ovo je ručna finansijska intervencija i mora ići kroz tehničkog + finansijskog owner-a.

---

## 10. Posebni recovery scenariji

### 10.1. PDF je u `APP_BANKA_ERROR`

Simptom:

* PDF nije obrađen;
* nema `BIM-*` redova;
* PDF je pomeren u error folder.

Postupak:

1. Otvori PDF i proveri da li ima tekst ili je sken/slika.
2. Proveri da li sistem može da izvuče `BrojIzvoda`, `DatumIzvoda`, `BrojRacuna`.
3. Ako je format banke promenjen, ne pokušavati ručno masovni import.
4. Ako je greška privremena, vrati PDF u inbox i ponovo pokreni `ImportBankaInbox_TX`.
5. Ako je PDF validan, ali parser ne podržava format, tehnički owner popravlja parser.

### 10.2. PDF je u `APP_BANKA_PROCESSED`, ali redovi nisu u `tblBankaImport`

Ovo je nekonzistentno stanje između fajl sistema i Excel TX-a.

Postupak:

1. Ne ubacivati isti PDF ponovo bez provere.
2. Proveri `tblBankaImport` po `IzvorFajl`, `BrojDokumenta`, `BankaReferenz`.
3. Proveri log greške import-a.
4. Ako nema redova i nema duplikata, tehnički owner može vratiti PDF iz processed u inbox i ponoviti import.
5. Zabeležiti u incident ticket-u da je PDF ručno vraćen.

### 10.3. `BIM.Obradjeno = Da`, ali nema `NOV-*`

Ovo je opasno nekonzistentno stanje.

Postupak:

1. Ne menjati odmah `Obradjeno` nazad na prazno.
2. Tražiti `NOV-*` po `BIM:<id>` u `Napomena`.
3. Tražiti po datumu, iznosu, partneru, broju izvoda.
4. Proveriti `Journal/` i backup.
5. Ako se dokaže da `NOV-*` ne postoji, tehnički owner može vratiti `Obradjeno` u prazno i ponoviti mapiranje.
6. Ako postoji sumnja, ne mapirati ponovo jer rizikuješ dupli finansijski red.

### 10.4. Uplata kupca otišla u avans, a trebalo je na fakturu

Simptom:

* `NOV.Tip = NOV_KUPCI_AVANS`;
* `FakturaID` prazan;
* kupac ima otvorenu fakturu koja odgovara uplati.

Postupak:

1. Proveri `BIM.PozivNaBroj` i `SvrhaPlacanja`.
2. Proveri otvorene fakture kupca.
3. Ako je uplata jednoznačno za fakturu, finansijski owner odobrava vezivanje.
4. Povezati `NOV-*` na ispravan `FakturaID` ili izvršiti korekciju kroz postojeću finansijsku proceduru.
5. Pokrenuti `UpdateFakturaStatus(fakturaID)`.
6. Ako je avans delimično potrošen, ne dirati bez analize split logike.

### 10.5. Uplata je vezana na pogrešnu fakturu

Postupak:

1. Identifikuj pogrešni `NOV-*`.
2. Identifikuj pogrešni `FakturaID` i ispravni `FakturaID`.
3. Izračunaj status obe fakture pre korekcije.
4. Finansijski owner mora odobriti korekciju.
5. Promeni vezu ili napravi storno/korekcioni red po pravilima.
6. Pokreni `UpdateFakturaStatus` za obe fakture.
7. Dokumentuj razlog.

### 10.6. Isplata kooperantu otišla u avans, a trebalo je na otkup

Simptom:

* `NOV.Tip = NOV_VIRMAN_AVANS_KOOP`;
* `OtkupID` prazan;
* postoji otvoren otkup za kooperanta/blok.

Postupak:

1. Proveri `BIM.PozivNaBroj` kao broj bloka.
2. Proveri `GetOtkupCandidatesForKooperantBlock` rezultat ili ručno otvorene otkupe.
3. Ako se jednoznačno zna otkup, finansijski owner odobrava vezivanje.
4. Povezati `NOV-*` na `OtkupID` ili napraviti korekciju.
5. Pokrenuti `UpdateOtkupStatus(otkupID)`.

### 10.7. Isplata je raspoređena na pogrešan blok

Postupak:

1. Identifikuj sve `NOV-*` redove nastale iz istog `BIM-*` reda.
2. Proveri koji `OtkupID` je popunjen na svakom.
3. Proveri da li je postojao višak knjižen kao avans.
4. Finansijski owner odlučuje korekciju.
5. Ako se menja, ažurirati statuse svih pogođenih `OtkupID` redova.
6. Popraviti `PartnerMap` ako je automatski izbor kooperanta bio pogrešan.

### 10.8. Pogrešan `PartnerMap`

Simptom:

* isti bankarski naziv se stalno mapira na pogrešnog partnera;
* auto-preview pokazuje pogrešan `KupacID`, `KooperantID` ili `OMID`.

Postupak:

1. Zaustaviti auto-mapiranje za taj partner.
2. Proveriti red u `tblPartnerMap`.
3. Proveriti sve `NOV-*` redove nastale posle pogrešne mape.
4. Tehnički owner ispravlja `tblPartnerMap`.
5. Finansijski owner odlučuje šta sa već pogrešno knjiženim redovima.
6. Testirati auto-preview pre nastavka auto-mapiranja.

### 10.9. Isti izvod uvezen dva puta

Postupak:

1. Proveri `BrojDokumenta` i `IzvorFajl` u `tblBankaImport`.
2. Proveri `BankaReferenz` za redove.
3. Ako su dupli `BIM-*` redovi bez `NOV-*`, stornirati/isključiti višak staging redove uz ticket.
4. Ako su dupli `NOV-*` redovi, finansijski owner mora odobriti storno/korekciju.
5. Proveriti da li je uzrok promenjen naziv fajla, izostanak reference ili parse bug.

### 10.10. Dve realne identične uplate tretirane kao duplikat

Postupak:

1. Uporedi originalni PDF izvod.
2. Ako postoje dve realne transakcije, proveri da li obe imaju različit `BankaReferenz`.
3. Ako `BankaReferenz` nedostaje, trenutni dedupe može ih spojiti.
4. Tehnički owner mora ručno dodati propušteni `BIM-*`/`NOV-*` ili popraviti parser/dedupe.
5. Dokumentovati slučaj kao dedupe false positive.

---

## 11. Avans procedure

### 11.1. Kupac avans → faktura

`ApplyAvansToFaktura_TX(kupacID, fakturaID)` radi transakcijski:

* traži otvorene avans uplate kupca;
* proverava iznos fakture i već uplaćeno;
* ako avans pokriva ceo preostali iznos, linkuje postojeći avans red na fakturu;
* ako je avans veći od preostalog iznosa, smanjuje originalni avans red i pravi split `NOV-*` red za potrošeni deo;
* ažurira status fakture kada je pokrivena.

Runbook pravila:

```text
Ako je avans splitovan, nikada ne menjaš samo jedan red bez razumevanja original/split para.
Ako se korekcija radi posle split-a, moraš identifikovati originalni avans NOV-ID i split NOV-ID.
Ako je faktura već postala plaćena, moraš rerun statusa posle korekcije.
```

### 11.2. Kooperant avans → otkup

Kod bankarskog mapiranja odlazne isplate:

* ako postoje otvoreni otkupi za kooperanta/blok, isplata se raspoređuje na njih;
* ako iznos prelazi otvorene otkupe, višak ide kao kooperantski avans;
* ako nema kandidata, sve ide kao avans.

Runbook pravila:

```text
Ako je isplata delimično raspoređena, prati sve NOV-* redove iz istog BIM-* reda.
Ako menjaš blok, moraš ažurirati sve pogođene OtkupID statuse.
Ako je višak knjižen kao avans, poslovni owner odlučuje da li ostaje avans ili se veže na drugi otkup.
```

---

## 12. Kako sprečavaš dupli finansijski red

Sistem ima zaštite:

1. PDF import dedupe kroz `BrojDokumenta + BankaReferenz` ili fallback na datum/iznos/partner.
2. `ValidateBankaImportNotProcessed` blokira mapiranje već obrađenih staging redova.
3. `Obradjeno = Da` označava da staging red ne sme opet u `tblNovac`.
4. `Napomena` u `tblNovac` nosi BIM trag.
5. `AutoMap..._TX` i `Map..._TX` koriste transakcije i snapshot više tabela.
6. `PartnerMap` sprečava nasumično pogađanje partnera svaki put, ali mora biti tačan.

Operativno pravilo:

> Za isti `BIM-*` red sme postojati jedan finansijski ishod. Ako postoji više `NOV-*` redova iz istog `BIM-*`, to mora biti namerna raspodela ili split, nikad slučajni remap.

---

## 13. Admin/VBA komande

Koristiti samo ako UI nije dovoljan ili ako tehnički owner radi incident.

```vb
' Import svih PDF izvoda iz inbox-a
Call ImportBankaInbox_TX

' Import jednog PDF-a, tehnička analiza
Call ImportOnePdfIntoBankaImport("C:\path\izvod.pdf")

' Automatsko mapiranje jedne BIM stavke
Debug.Print AutoMapBankaImportRow_TX("BIM-00001")

' Automatsko mapiranje svih otvorenih BIM stavki
Debug.Print AutoMapAllBankaImport_TX()

' Ručno mapiranje kao kupac, bez fakture = avans
Debug.Print MapBankaImportAsKupac_TX("BIM-00001", "KUP-00001", "", True)

' Ručno mapiranje kao kupac na fakturu
Debug.Print MapBankaImportAsKupac_TX("BIM-00001", "KUP-00001", "FAK-00001", True)

' Ručno mapiranje kao kooperant po bloku iz poziva na broj
Debug.Print MapBankaImportAsKooperantBlock_TX("BIM-00001", "KOO-00001", True)

' Ručno mapiranje kao kooperant po ručno zadatom bloku
Debug.Print MapBankaImportAsKooperantBlockManual_TX("BIM-00001", "KOO-00001", "BLOK-123", True)

' Ručno mapiranje kao OM / stanica
Debug.Print MapBankaImportAsOM_TX("BIM-00001", "ST-00001", "", True)

' Preskoči BIM stavku
Debug.Print SkipBankaImportRow_TX("BIM-00001")

' Primeni kupčev avans na fakturu
Debug.Print ApplyAvansToFaktura_TX("KUP-00001", "FAK-00001")

' Recalculate statusi posle korekcije
Call UpdateFakturaStatus("FAK-00001")
Call UpdateOtkupStatus("OTK-00001")
```

Ne koristiti direktno `SaveNovac` za incident recovery osim ako tehnički i finansijski owner eksplicitno odluče korekcioni finansijski unos.

---

## 14. Ko donosi odluku

### Operator sme sam

* importovati PDF iz inbox-a;
* koristiti auto-map kada preview jasno odgovara očekivanju;
* ručno mapirati otvorenu `BIM-*` stavku kada je partner/faktura/otkup jednoznačan;
* označiti `Skip` ako postoji jasna operativna politika za takve stavke;
* osvežiti formu i proveriti stanje.

### Tehnički owner odlučuje

* vraćanje PDF-a iz processed/error u inbox;
* vraćanje `BIM.Obradjeno` iz `Da`, `Skip` ili `Error` nazad u prazno;
* ručnu izmenu `tblPartnerMap`;
* ručnu korekciju `tblNovac` strukture;
* recovery kada `BIM.Obradjeno = Da`, ali nema `NOV-*`;
* popravku parsera/dedupe logike;
* slučajeve gde TX rollback i file move nisu konzistentni.

### Finansijski / poslovni owner odlučuje

* da li uplata ide na fakturu ili u avans;
* na koju fakturu ide uplata ako ih ima više otvorenih;
* na koji otkup/blok ide isplata kooperantu;
* da li se višak isplate vodi kao avans;
* da li se pogrešni `NOV-*` red stornira ili koriguje;
* da li je `Skip` opravdan;
* šta raditi kod duplog ili propuštenog bankarskog reda.

### Niko ne sme bez odobrenja

* brisati `NOV-*` redove;
* brisati `BIM-*` redove;
* ručno menjati `Obradjeno = Da` nazad u prazno bez provere duplikata;
* mapirati `BIM-*` drugi put zato što “ne vidi uplatu”; prvo se traži postojeći `NOV-*`;
* menjati `PartnerMap` bez provere postojećih posledica;
* menjati `FakturaID`/`OtkupID` na finansijskim redovima bez recalculation statusa.

---

## 15. Checklist za zatvaranje incidenta

```text
[ ] Identifikovan BIM-ID
[ ] Proveren IzvorFajl / PDF lokacija
[ ] Proveren BrojDokumenta i BankaReferenz
[ ] Proveren Obradjeno status
[ ] Proveren da li postoji NOV-ID
[ ] Proveren BIM trag u NOV.Napomena
[ ] Proveren PartnerMap ako je auto-map korišćen
[ ] Ako je kupac, proveren KupacID i FakturaID / avans
[ ] Ako je kooperant, proveren KooperantID i OtkupID / blok / avans
[ ] Ako je OM, proveren OMID
[ ] Ako je korekcija, postoji odluka finansijskog owner-a
[ ] Ako je ručna tehnička intervencija, postoji backup/ticket
[ ] Pokrenut UpdateFakturaStatus za pogođene fakture
[ ] Pokrenut UpdateOtkupStatus za pogođene otkupe
[ ] Korisnik obavešten
```

---

## 16. Primeri odluke

### Primer A: PDF u processed, `BIM-*` postoji, `Obradjeno = prazno`

Zaključak: import je uspeo, stavka čeka mapiranje.
Akcija: otvoriti `frmBankaImport`, proveriti preview, mapirati auto ili ručno.

### Primer B: Kupac uplatio tačan iznos fakture, ali otišlo u avans

Zaključak: faktura nije jednoznačno resolvovana tokom mapiranja.
Akcija: finansijski owner odobrava vezivanje avansa na fakturu; zatim update status fakture.

### Primer C: `BIM.Obradjeno = Da`, korisnik kaže “nema uplate”

Zaključak: prvo pronaći `NOV-*`; možda je uplata u avansu ili na pogrešnoj fakturi.
Akcija: ne remapirati. Tražiti po `BIM:<id>` u napomeni, iznosu i datumu.

### Primer D: `PartnerMap` mapira bankarski naziv na pogrešnog kupca

Zaključak: sistemski rizik za sve buduće izvode.
Akcija: zaustaviti auto-map, ispraviti mapu, pregledati sve pogođene `NOV-*` redove.

### Primer E: Isplata kooperantu veća od otvorenih otkupa

Zaključak: deo ide na otvorene otkupe, višak kao avans.
Akcija: proveriti sve `NOV-*` redove iz istog `BIM-*`, ne menjati samo jedan split deo.

### Primer F: Dve iste uplate istog dana od istog partnera, jedna fali

Zaključak: moguć false duplicate ako nema `BankaReferenz`.
Akcija: proveriti original PDF i bankarsku referencu; tehnički owner rešava ručni dodatak ili dedupe fix.

---

## 17. Poznate production rupe koje treba zatvoriti

1. Dodati `BIM-ID` kao strukturisanu kolonu u `tblNovac`, ne samo u `Napomena`.
2. Dodati `tblBankaEventLog`: import, map, skip, error, remap, correction, operator, razlog.
3. Dodati eksplicitnu “Undo/Correct bank mapping” proceduru umesto ručnih korekcija.
4. Dodati obavezan razlog za `Skip`.
5. Dodati audit za izmene `tblPartnerMap`.
6. Dodati ekran “Pronađi NOV po BIM” direktno u `frmBankaImport`.
7. Dodati detekciju `BIM.Obradjeno = Da`, ali bez `NOV-*` linka.
8. Dodati dnevni report `BIM.Obradjeno = Error` i starih otvorenih `BIM` stavki.
9. Dodati bolju zaštitu od file-system/TX mismatch-a kod pomeranja PDF-a.
10. Dodati formalnu politiku za bankarske naknade, interne transfere i `Skip`.
11. Dodati reconciliation report: `tblBankaImport` vs `tblNovac` vs `tblFakture`/`tblOtkup` status.
12. Dodati test slučajeve za dve identične transakcije bez `BankaReferenz`.

Do tada važi konzervativno pravilo:

> Bankarska stavka prvo živi kao `BIM-*`. Finansijski efekat postoji tek kada postoji validan `NOV-*`. Ako ne znaš vezu `BIM-* → NOV-* → Faktura/Otkup`, sistem nije u konzistentnom finansijskom stanju.
