# Nalaz review-a trenutnog `main` — 2026-10-02

Repo: `dusanmiladinovicVNM/otkupapp-pwa`

Bazni nalaz iz prethodnog chata bio je na:

```text
1095fddc419c9d6a1b12b64eea43273fcac90c1f
PR #393 — S5-5a: JS kapija za PWA/GAS sloj
```

Trenutni `main` u ovom nalazu je:

```text
25e482c1df45809f2198f913c65fc88c00139f92
PR #400 — AMB-10b-1: knjiga dobija pisca, bez cutovera
```

Diff od prethodnog nalaza do trenutnog `main`:

```text
status: ahead
ahead_by: 56
behind_by: 0
total_commits: 56
```

## Kratak verdict

Prethodni centralni rizik bio je:

```text
S5-5b — PWA/GAS otkup kao header + stavke
```

Taj rizik je sada zatvoren.

Novi centralni rizik je:

```text
AMB-10b-2 — cutover ambalazne knjige
```

Ocena nivoa projekta:

```text
prethodno: 4.25 / 5
sada:      4.35 / 5
```

Ne više od toga, jer je AMB-10 trenutno veliki otvoreni tranzicioni rez: ugovor i pisac su ozbiljni, ali cutover još nije završen.

---

## Šta je zatvoreno u odnosu na prethodni nalaz

### 1. S5-5b više nije otvoren TODO

`docs/STANJE_REFAKTORA.md` kaže da je **S5 zatvoren kroz #385–#394**, da je **S5-5b spojen**, i da sada ostaje samo jedno mesto `DEGRADIRANO` grane.

To menja status prethodnog P0 nalaza:

```text
S5-5b: završiti PWA/GAS otkup kao header + stavke
```

u:

```text
ZATVORENO
```

### 2. PWA lokalni model sada nosi `stavke[]`

Novi fajl:

```text
src/js/features/otkup/otkup-stavke.js
```

uvodi jedno mesto koje zna šta je stavka otkupa. Komentar je važan:

```text
Zapis otkupa od ovog reza NE nosi klasu, kolicinu, cenu i ambalazu na sebi:
nosi stavke[], red po klasi.
```

Scope je dobar i čist:

```text
UI u ovom rezu i dalje unosi jednu klasu.
Model je već spreman za N > 1.
Čitaoci sabiraju preko helpera, ne sami.
```

To je ispravno: žica i model se rešavaju pre UX proširenja.

### 3. JS harness sada meri S5-5b

`tests/js/pokreni.js` sada uključuje:

```text
require('./suites/gas-otk-stavke')
require('./suites/otkup-stavke')
```

To znači da S5-5b nije samo napisan, nego i uključen u dokazni harness.

Najbitnije tvrdnje u `tests/js/suites/gas-otk-stavke.js` pokrivaju:

```text
- otkup bez stavki se odbija po imenu
- stavka bez svog ClientRecordID-a se odbija
- dve stavke iste klase se odbijaju
- dve stavke sa istim ClientRecordID-em se odbijaju
- iste stavke u drugom redosledu nisu razlika
- izmenjena količina jeste razlika
- partial upis sa izmenjenim sadržajem jeste konflikt čak i bez zaglavlja
- isti item CRID pod drugim otkupom jeste konflikt
- završen dokument ne prima novu stavku
- doslovno isti završen dokument je idempotentan
- ugovor kolona je doslovno isti kao u VBA
- indeks taba pada po imenu kao VBA čitalac
```

Ovo je tačno klasa problema koja je pre bila samo logički pročitana, a sada je mereno.

### 4. VBA master sync ima stvarni `OTK_STAVKE` ugovor

`modMasterSync` sada centralizuje:

```text
OtkZaglavljeKolone()
OtkStavkeKolone()
```

Bitna semantika:

```text
Zaglavlje od S5-5b nosi samo činjenice zaglavlja.
Klasa/Kolicina/Cena/KolAmbalaze su mrtvi slotovi koje oba pisca ostavljaju prazne.
StavkeCount je manifest.
OTK_STAVKE je red po stavci.
```

`OtkPwaStavkeIzTaba` je fail-closed:

```text
- red sa OtkupClientRecordID mora imati svoj ClientRecordID
- dupli item ClientRecordID obara čitanje
- desktop push redovi se preskaču jer nemaju OtkupClientRecordID
- čitanje OTK_STAVKE taba ne sme tiho postati prazan skup ako je mreža/JSON pao
```

Ovo je veliki napredak u odnosu na prethodni nalaz.

---

## Novi centralni front: AMB-10

Trenutni `main` je sada dominantno AMB-10 rad.

### 1. Novi domenski ugovor

Dodat je:

```text
src-vba/modAmbalazaUgovor.bas
```

Ugovor definiše:

```text
- append-only knjigu prenosa
- nalog kao stranu prenosa: OdNalog -> NaNalog
- zatvorene vrste naloga
- zatvorene vrste kretanja
- sistemske naloge: Firma, SpoljniSvet
- partner/sopstveni/granica klase naloga
- fail-closed razrešavanje naloga
- doprinos obavezi
```

Ovo je dobar smer: ambalaža više nije relativni `Ulaz/Izlaz` zapis po entitetu, nego događaj prenosa sa obe strane.

### 2. Pisac knjige postoji

`modAmbalaza` sada ima:

```text
PrenesiAmbalazu(...)
```

Komentar definiše važnu stvar: jedan poziv može da upiše do tri reda:

```text
pokrice deficita   SpoljniSvet -> izvor   ULAZ_TUDJE_AMBALAZE
trazeni prenos     od -> na               trazena vrsta
ostatak podele     od -> na               IZDATA_PRAZNA
```

To sprovodi ključne invarijante:

```text
AMB-INV-07 — nijedan realan nalog ne sme posle commit-a biti ispod nule
AMB-INV-09 — obaveza partneru ne sme postati negativna
AMB-INV-10 — dokument zaključava jedan neuređen par {Od, Na}
AMB-INV-11 — AmbID jedinstven, StornoOd pokazuje na tačno jedan red
```

### 3. Pisac je dobar, ali cutover nije gotov

PR #400 je nazvan:

```text
AMB-10b-1 — knjiga dobija pisca, bez cutovera
```

Dakle status nije:

```text
AMB-10 završeno
```

nego:

```text
AMB-10 ugovor + dokument + pisac postoje;
cutover još čeka.
```

To je najvažniji novi nalaz.

---

## Ažurirana TODO lista

## P0 — trenutno najvažnije

### 1. AMB-10b-2 cutover

Sledeći glavni posao treba da bude prebacivanje živih puteva na novu ambalažnu knjigu.

Trenutno stanje:

```text
model        ✅
10a ugovor   ✅
10-DOK       ✅
10b-1 pisac  ✅
10b-2 cutover ⏳
```

Cutover mora dokazati da živi poslovni putevi ne nastavljaju da zavise od starog relativnog `Ulaz/Izlaz` modela.

### 2. AMB-10 acceptance pre cutover merge-a

Minimalni acceptance za cutover:

```text
- stari i novi čitaoci ne daju različite salde u podržanim scenarijima
- deficit potvrda radi samo gde sme
- ULAZ_TUDJE_AMBALAZE ne sme biti eksplicitan zahtev
- negativna obaveza partneru pada
- dokument zaključava jedan neuređen par {Od, Na}
- dupli AmbID pada
- StornoOd mora pokazivati na tačno jedan postojeći red
- jedan zahtev koji generiše više redova ostaje idempotentan po zbiru, ne po jednom redu
```

### 3. S6 prijemnica ostaje parkirana do posle AMB-10

`docs/STANJE_REFAKTORA.md` kaže da je S6 parkiran na koraku 1/8 i nastavlja se posle AMB-10.

Dakle redosled sada nije:

```text
S5-5b -> E2E -> S6
```

nego:

```text
AMB-10b-2 cutover
-> AMB-10 storno/kontra-stav ako je sledeći planirani rez
-> S6 prijemnica
```

---

## P1 — i dalje aktuelno

### 4. Verzionisanje

I dalje nije rešeno kao sistem.

Potrebno razdvojiti:

```text
APP_VERSION
release tag / git tag
BUILD_VERSION
v6-ui-* build
S5/S6/AMB roadmap faze
architecture changelog version
```

Ovo nije blocker za sledeći AMB rez, ali jeste blocker za ozbiljan rollout.

### 5. UI registry/static checker

I dalje aktuelno.

Potrebna kapija za:

```text
- registry red ima modul ili validan panel
- modul postoji
- Scr_Meta postoji
- Scr_Dozvoljen postoji ako registry kaže da postoji
- Scr_Rows / Scr_Liste / Scr_Radnje format je validan
- nema mrtvog ekrana u registru
- nema ekranskog modula koji nije povezan
```

### 6. CI status / branch protection ručno proveriti

GitHub connector u review-u nije vratio workflow run/status za trenutni commit.

To ne znači da CI ne radi, nego da nije potvrđen kroz connector.

Pre oslanjanja na `main` kao zelen, proveriti ručno:

```text
GitHub Actions tab
zadnji static run
branch protection
required checks
push to main trigger
PR trigger
```

### 7. Release checklist formalizovati

Checklist sada mora uključiti i AMB-10 smoke:

```text
static CI
JS syntax
JS harness
JS harness self-test
Excel compile
VBA targeted suite
VBA full suite po potrebi
AMB-10 targeted suite
ručni smoke ambalaže
release publish
backup/rollback
```

---

## P2 — ostaje otvoreno, ali nije sledeći rez

### 8. Bank parser golden fixtures

I dalje korisno:

```text
Komercijalna / NLB
Halk
Alta
ProCredit
```

Za svaku banku treba sample text fixture i expected parsed result.

Ali trenutno nije najkritičniji front; AMB-10 cutover je veći rizik.

### 9. Politika izdatih dokumenata

I dalje otvoreno.

Potrebno jasno definisati:

```text
kada sme in-place korekcija
kada mora ispravka/revizija
šta se štampa
šta se vidi u UI
šta ostaje u audit tragu
```

### 10. Architecture reference / changelog update

`docs/STANJE_REFAKTORA.md` je sada bolji kratki source of truth od starog architecture changelog-a.

Potrebno je vremenom uskladiti:

```text
ARCHITECTURE_REFERENCE
ARCHITECTURE_CHANGELOG
STANJE_REFAKTORA
DOMEN dokumente
```

---

## Skinuto sa prethodne TODO liste

Više ne voditi kao otvoreno:

```text
PWA/GAS nema automatsku proveru
JS sintaksa nije proverena
JS harness nema dokaz da pocrveni
IndexedDB atomic claim je samo pročitan
S5-5b PWA otkup header + stavke nije završen
GAS/VBA ugovor OTK_STAVKE nije meren
partial upis/retry/item CRID kolizije nisu pokrivene
```

Sve to je ili zatvoreno ili značajno spušteno po riziku.

---

## Sledeći preporučeni redosled

```text
1. AMB-10b-2 cutover
2. AMB-10 storno/kontra-stav tok, ako je sledeći planirani rez
3. AMB acceptance suite / smoke na realnim scenarijima
4. S6 prijemnica
5. verzionisanje + release checklist
6. UI registry checker
7. bank parser fixtures
```

## Konačni sažetak

Projekat je napredovao od prethodnog nalaza.

S5 više nije centralni rizik. PWA/GAS otkup header+stavke je zatvoren kroz model, GAS/VBA ugovor i JS harness.

Centralni rizik je sada AMB-10: ambalažna knjiga ima dobar ugovor i pisca, ali sistem je još u tranziciji jer cutover nije završen.

Najkraći verdict:

```text
Nalaz se poboljšao.
Glavni otvoreni rizik se pomerio sa S5/PWA na AMB-10/cutover.
```