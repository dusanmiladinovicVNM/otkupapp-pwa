# Infosys → AgriX migration discovery checklist

**Svrha:** pre ponude utvrditi da li je migracija tehnički i operativno bezbedna i koji podaci, procesi i rizici postoje.

## 1. Sistem i instalacija

- naziv i verzija Infosys rešenja;
- korišćeni moduli;
- lokalna, mrežna ili cloud instalacija;
- baza podataka ili fajlovi koje sistem koristi;
- broj korisnika i radnih stanica;
- broj firmi, lokacija i sezonskih terminala;
- ko administrira sistem;
- ugovor o održavanju i podršci;
- period i uslovi raskida postojećeg ugovora.

## 2. Podaci dostupni za izvoz

- partneri i kooperanti;
- poljoprivredna gazdinstva;
- otkupna mesta i stanice;
- artikli, kulture, sorte i klase;
- cenovnici i obračunska pravila;
- prijemnice;
- otkupni listovi;
- isplate i obaveze;
- reversi i ambalaža;
- laboratorijski rezultati;
- skladišni dokumenti;
- fakture i SEF veze;
- bankovni nalozi i izvodi;
- istorija izmena i audit podaci;
- korisnici, uloge i ovlašćenja.

Za svaki skup podataka zabeležiti:

- format: CSV / XLSX / XML / SQL / PDF / štampa / nedostupno;
- broj redova i vremenski opseg;
- jedinstveni identifikator;
- kvalitet i duplikate;
- vlasništvo nad podacima;
- pravo na izvoz;
- potrebu za čišćenjem ili ručnim mapiranjem.

## 3. Integracije

- vaga i serijski/LAN protokol;
- laboratorijska oprema;
- štampači i obrasci;
- fiskalizacija;
- računovodstveni program;
- BizniSoft ili drugi ERP;
- SEF;
- banka i platni nalozi;
- Google Drive/Sheets;
- email;
- mobilni uređaji;
- druge interne aplikacije.

## 4. Procesna mapa

Za svaki ključni proces dokumentovati:

1. ko pokreće proces;
2. koji podaci ulaze;
3. koji dokument nastaje;
4. ko proverava i odobrava;
5. gde se podatak dalje koristi;
6. šta se radi kada nema interneta ili sistem ne radi;
7. sezonski maksimum operacija na sat i dnevno.

Obavezni procesi:

- prijem proizvođača/robe;
- merenje;
- klasiranje i kvalitet;
- obračun cene;
- dokumenti otkupa;
- isplata;
- ambalaža;
- skladište;
- transport;
- finansijska kontrola;
- korekcije, storno i reklamacije;
- zatvaranje dana i sezone.

## 5. Migracioni rizici

- nema potpunog izvoza;
- različiti identifikatori istog partnera;
- duplikati i neaktivni partneri;
- istorijski dokumenti bez veza;
- formule ili obračuni koji nisu dokumentovani;
- ručne Excel evidencije van sistema;
- custom funkcije poznate samo jednom zaposlenom;
- zavisnost od starog dobavljača za izvoz;
- paralelna sezona i premalo vremena za obuku;
- loša mreža na stanicama;
- neusklađeni štampani obrasci;
- računovodstvene razlike;
- neprihvaćena odgovornost za početna stanja.

## 6. Plan prelaska

- datum zamrzavanja konfiguracije;
- datum probnog izvoza;
- testna migracija;
- validacija uzoraka;
- korisničko prihvatanje;
- obuka ključnih korisnika;
- paralelni rad kada je potreban;
- finalni izvoz;
- cutover datum;
- rollback kriterijumi;
- pojačana podrška u prvim danima;
- kontrola posle prvog obračuna i prve isplate.

## 7. Acceptance kriterijumi

Migracija se ne smatra završenom dok nisu potvrđeni:

- broj prenetih partnera;
- broj prenetih otvorenih stavki;
- početna stanja ambalaže i skladišta;
- uzorci istorijskih dokumenata;
- obračun najmanje tri reprezentativna scenarija;
- štampa svih obaveznih dokumenata;
- prava korisnika;
- backup i recovery;
- monitoring;
- potpis odgovorne osobe klijenta na migracioni zapisnik.

## 8. Go / no-go odluka

### Go

- podaci su dostupni i razumljivi;
- ključni obračuni su potvrđeni;
- korisnici su dostupni za test;
- postoji realan period pre sezone;
- rollback je moguć;
- vlasnik prihvata migracioni plan i odgovornosti.

### No-go / odlaganje

- nema pristupa podacima;
- ključna pravila nisu poznata;
- migracija se planira neposredno pred peak sezone;
- nema odgovorne osobe za validaciju;
- zahtev je istovremeno migracija i veliki nedovršeni custom razvoj;
- rollback nije moguć;
- klijent ne prihvata proveru početnih stanja.
