# Ambalaža — model podataka

> Holistički pregled, 27.09.2026. Nastao je pred rez „povrat ambalaže živi samo u
> `tblAmbalaza`", kad se pokazalo da ambalaža dodiruje **otkup, otpremnicu,
> prijemnicu, izlaz kupcima i revers** — pa rez nad dve kolone ne može da se
> donese bez modela.
>
> Sve u ovom dokumentu je **mereno u kodu**, sa `fajl:linija`. Gde nešto nije
> mereno, tako i piše.

## 1) Šta `tblAmbalaza` jeste

**Jednostran, entitetski-relativan knjigovodstveni zapis.** Jedan red = jedno
kretanje **iz ugla jednog entiteta**: `EntitetID` + `EntitetTip` + `Smer`
(`Ulaz`/`Izlaz`) + `Kolicina` + `TipAmbalaze`.

Model je namerno ovakav i zapisan je u kodu ([modAmbalaza.bas:47](../../src-vba/modAmbalaza.bas)):

> „Ambalaza se upisuje JEDNOM, entitetski-relativno. Vozac je transporter
> = INVERZNI protivpartner entiteta."

`VozacAmbEffectiveSmer(smer, entitetTip)` izvodi vozačevu stranu umesto da je
upisuje: šta entitetu uđe, iz vozača izlazi. Kompletna ruta daje **saldo 0**, a
otvorena otpremnica ostaje kao pozitivan saldo kod vozača — to je i svrha.

**`tblAmbalaza` nije duplikat nijednog dokumenta.** Saldo se već računa isključivo
iz nje: `GetAmbalazeStanje`, `GetStanicaAmbSaldo`, `GetVozacAmbSaldo`,
`GetKooperantAmbOpening`. Nijedan bilans ne čita količinu sa dokumenta.

## 2) Ko je nosilac salda — pravilo koje objašnjava svih devet knjiženja

| Uloga | Nosilac salda? | Kako se vodi |
|---|---|---|
| **Kooperant** | da | svoj red |
| **Stanica (OM)** | da | svoj red |
| **Kupac (hladnjača)** | da | svoj red |
| **Vozač** | da, ali **izveden** | `VozacID` na redu + `VozacAmbEffectiveSmer` |
| **Firma (centrala)** | **ne** | spoljni svet; noga se ne upisuje |

> **AMB-01.** Red se upisuje za **svakog nosioca salda** koji u kretanju
> učestvuje. Protivpartner dobija svoj red **samo ako je i sam nosilac salda**;
> ako je transporter — izvodi se; ako je spoljni svet — ne vodi se.

Ovo pravilo objašnjava **sve** izmerene slučajeve, uključujući one koji na prvi
pogled izgledaju nedosledno (dve noge kod otkupa, jedna kod otpremnice).

## 3) Matrica knjiženja — mereno

| # | Gde | Količina | Smer · Entitet | Vozač | `DokumentID` | `DokumentTIP` |
|---|---|---|---|---|---|---|
| 1 | [modOtkup:1445](../../src-vba/modOtkup.bas) | zbir stavki | Izlaz · Kooperant | — | `OtkupID` | `Otkup` |
| 2 | modOtkup:1448 | zbir stavki | Ulaz · Stanica | — | `OtkupID` | `Otkup` |
| 3 | modOtkup:1457 | `KolAmbIzdata` | Ulaz · Kooperant | — | `OtkupID` | **`OM-Izlaz-Koop`** |
| 4 | modOtkup:1460 | `KolAmbIzdata` | Izlaz · Stanica | — | `OtkupID` | **`OM-Izlaz-Koop`** |
| 5 | [modDokumenta:4969](../../src-vba/modDokumenta.bas) | ukupno | Izlaz · Stanica | da | `OtpremnicaID` | `Otpremnica` |
| 6 | modDokumenta:6857 | `KolAmbalaze` | Ulaz · Kupac | da | `PrijemnicaID` | `Prijemnica` |
| 7 | modDokumenta:6863 | `KolAmbVracena` | Izlaz · Kupac | da | `PrijemnicaID` | **`Prijemnica`** |
| 8 | modDokumenta:7026 | `kolAmb` | Izlaz · Kupac | da | **`brojDok`** | `Kupci-Otpremnica` |
| 9 | modDokumenta:8619–8669 | revers, 4 smera | 2 noge (KOOP) / 1 noga (FIRMA) | samo FIRMA | **`brojDok`** | `OM-*` |

Redovi 1–2 i 3–4 su dvonožni jer su **oba** učesnika nosioci salda i vozača nema.
Red 5–8 su jednonožni jer je protivpartner **vozač**. Red 9 je oba oblika, po
smeru: KOOP je dvonožan, FIRMA jednonožan jer firma nije nosilac salda.

## 4) Tri napetosti — i jedna stvar koja NIJE problem

### T1 — `DokumentID` je ponegde ID, ponegde **broj** ⚠ IDENTITY RISK

| Put | Identitet dokumenta | Šta stvarno stoji u `DokumentID` |
|---|---|---|
| otkup · otpremnica · prijemnica | `OtkupID` / `OtpremnicaID` / `PrijemnicaID` | **ID** |
| revers (`OM-*`) | **`ReversID`** (REV-IDENT-01) | **broj** (labela) |
| izlaz kupcima | — | **broj** |

Dva identiteta u dve kolone, zamenjena između dva puta: kod dokumenata identitet
je u `DokumentID` a `ReversID` je prazan; kod reversa identitet je u `ReversID` a
`DokumentID` nosi **labelu**. Posledica je merljiva — storno ide **dvema**
putanjama: otkup hvata svoje noge po `DokumentID`
([modOtkup:1456](../../src-vba/modOtkup.bas) „Isti DokumentID -> storno otkupa
hvata i ovu nogu"), a revers preko `ReversIDRazresi`
([modStorno:1907](../../src-vba/modStorno.bas)).

Ovo je isti obrazac koji je `ZBR-IDENT-01` već jednom rešio za zbirnu: **broj je
labela, ne identitet.**

> **AMB-02 (predlog).** `DokumentID` nosi **isključivo identitet** izvornog
> dokumenta. Gde identiteta nema, ne upisuje se broj — dodaje se identitet.

### T2 — količina kretanja keširana na dokumentu

| Kolona | Pozicija | Već u knjizi |
|---|---|---|
| `tblOtkup.KolAmbIzdata` | 17 / 29 | redovi 3–4 |
| `tblPrijemnica.KolAmbVracena` | 13 / 28 | red 7 |

Obe su keš izvedene činjenice — isti obrazac kao `Isplaceno`/`DatumIsplate`
(obrisani u S1d, istina u `tblNovac`) i `Fakturisano`/`FakturaID` (odluka za S6).
Jedine su takve u celom kanonu: pretraga po „Vracen" daje **samo**
`tblPrijemnica.KolAmbVracena`.

> **AMB-03 (predlog).** Dokument ne nosi količinu kretanja ambalaže. Čitaoci idu
> na knjigu kroz jedan pomoćnik `AmbalazaKolicina(dokumentID, dokTip, smer)`.
>
> **Otvoreno poslovno pitanje:** knjiga svoje noge markira `Stornirano`, a kolona
> na dokumentu zadrži vrednost. Štampa storniranog dokumenta tada pokazuje ili
> **ono što je tada uneto**, ili **ono što važi danas** (nula). Dok se to ne
> odluči, čitalac se ne sme pisati.

### T3 — pozajmljen tip dokumenta pravi izuzetak u kapiji

Otkup knjiži svoje kretanje pod **`OM-Izlaz-Koop`** — tipom *revers dokumenta*.
Zato integritetska provera mora da pravi izuzetak
([modIntegritet:546](../../src-vba/modIntegritet.bas)):

```vb
' OtkupID-evi: ambalaza uz otkup ima tip reversa, ali nije revers
If modStorno.ReversTipJe(tip) And Not otkupi.Exists(dok) Then
```

Izuzetak je **tačan** i namerno napisan — ali postoji samo zato što je tip
pozajmljen. Kretanje uz dokument sa sopstvenim tipom ukinulo bi ga: pravilo
umesto izuzetka.

A prijemnica je u trećem svetu: povrat praznih knjiži pod `Prijemnica` (red 7),
dakle **van `OM-*` taksonomije** — u izveštajima ambalaže se ne vidi kao povrat,
iako to jeste.

> **AMB-04 (predlog, poslovna odluka).** Kretanje ambalaže uz dokument dobija
> svoj tip, odvojen od tipa revers dokumenta, i prijemnica ulazi u istu
> taksonomiju. Menja **šta izveštaji broje kao „izdate prazne"** — zato odluka
> nije tehnička. Dira 6 PROD mesta (`modBrojevi`, `modIntegritet`, tri izveštaja,
> `modPrint`).

### Šta NIJE problem — da se ne „popravlja"

- **Jednostranost knjige je namerna**, ne propust. Vozačeva strana se izvodi i
  tako ruta zatvara saldo na 0.
- **`Stavke.KolAmbalaze` nije keš.** Knjiga dobija **zbir**
  (`ZbirAmbalazeStavki(stavke)`, [modOtkup:785](../../src-vba/modOtkup.bas)), a
  linija nosi svoj broj — iz knjige se ne može rekonstruisati po klasi.
- **Saldo već čita knjigu.** Nijedan bilans ne zavisi od keša sa dokumenta, pa je
  rizik uklanjanja keša ograničen na štampu i izveštaje samog dokumenta.
- **Revers bez `ReversID`-a uz otkup nije rupa** — `Chk_B10` to eksplicitno
  isključuje, jer „ambalaza uz otkup ima tip reversa, ali nije revers".

## 6) Ciljni model — AMB-10 (v2, posle dizajn review-a)

> Odluke operatera 27.09.2026: **stampa storniranog dokumenta prikazuje ono sto je
> vazilo pre storna**; i pitanje da li ideal trazi dve noge u svakom dogadjaju.
>
> **Ne trazi.** Dve noge nisu cilj nego **simptom** — posledica reda koji ume da
> imenuje samo jednu stranu. Danas isti model zato proizvodi dva oblika: dva reda
> kad su obe strane nosioci salda, jedan red plus **pravilo** kad je druga strana
> vozac. Ideal ne dodaje drugu nogu — **ukida pojam noge.**
>
> **v2** je rezultat dizajn review-a: pravac je potvrdjen, specifikacija nije bila
> spremna za kod. Sta je review promenio stoji u 6.9 — ukljucujuci dve moje greske.

### 6.1 Dogadjaj je jedan red koji imenuje obe strane

```
tblAmbalaza  (dogadjaj = PRENOS, knjiga je APPEND-ONLY)

  AmbID              identitet dogadjaja
  Datum
  TipAmbalaze
  Kolicina           UVEK pozitivna -- smera nema, smer je (Od -> Na)

  OdNalogTip         Kooperant | Stanica | Kupac | Vozac | Firma | SpoljniSvet
  OdNalogID
  NaNalogTip         isto
  NaNalogID

  DokumentTIP        povod
  DokumentID         UVEK identitet, nikad broj
  VrstaKretanja      KOJI poslovni efekat dokumenta je ovo

  StornoOd           AmbID originala koji ovaj ponistava; inace prazno

  CreatedAt
  CreatedBy
```

**Nema `ModifiedAt`/`ModifiedBy`** — i to nije previd nego ugovor: knjiga se ne
menja. Kolona `ModifiedAt` nad append-only tabelom je poziv sledecem citaocu da
zakljuci da red sme da se prepise. Pogresan unos se ne popravlja UPDATE-om nego
**stornom i novim dogadjajem**.

**Sta nestaje, ne menja se nego nestaje:** `Smer` · `EntitetID`/`EntitetTip` ·
`VozacID` · `ReversID` · `Stornirano` · cela funkcija `VozacAmbEffectiveSmer` ·
izuzetak u `Chk_B10` · obe kes kolone (`tblOtkup.KolAmbIzdata`,
`tblPrijemnica.KolAmbVracena`) · druga storno putanja.

### 6.2 Zasto prenos, a ne dve noge

| | dve noge (klasicno dvojno) | **prenos (Od -> Na)** |
|---|---|---|
| „tacno dve strane" | trazi **kapiju** koja proverava da noge postoje i da se slazu — to je `Chk_B10` i danas | **strukturno nemoguce prekrsiti**: siroce noge nema jer noge nema |
| uparivanje | trazi identitet dogadjaja na svakom redu | red **jeste** dogadjaj |
| broj redova | otkup: 4 reda za 2 dogadjaja | otkup: **2 reda** |
| smer | kolona koja se moze pogresno upisati | **izveden iz para**, ne postoji kao podatak |

Ovo nije nov model nego **dovrsen postojeci**: danasnji red vec *misli* obe
strane — jednu upisuje, drugu podrazumeva. Ideal je prestati podrazumevati.

### 6.3 `VrstaKretanja` — sta dokument radi, odvojeno od toga koji je dokument

Danas tip dokumenta nosi **dva** posla: koji je dokument povod, i koje je to
kretanje. Zato otkup knjizi pod `OM-Izlaz-Koop` (tipom revers dokumenta) i zato
`Chk_B10` mora izuzetak. Razdvajanje ih resava oba:

```
DokumentTIP / DokumentID   -- KOJI dokument je povod
VrstaKretanja              -- KOJI njegov efekat je ovaj red
```

Izvedeno iz devet izmerenih mesta (§3); spisak je **zatvoren enum** i traži
potvrdu operatera pre koda:

| `VrstaKretanja` | Od -> Na | Danas |
|---|---|---|
| `ROBA_PRIMLJENA` | Kooperant -> Stanica | otkup, redovi 1–2 |
| `PRAZNA_IZDATA` | Stanica -> Kooperant | otkup 3–4, revers `IZDAVANJE` |
| `PRAZNA_PRIMLJENA` | Kooperant -> Stanica | revers `PRIJEM` |
| `ROBA_UTOVARENA` | Stanica -> Vozac | otpremnica |
| `ROBA_ISPORUCENA` | Vozac -> Kupac | prijemnica (pune) |
| `PRAZNA_VRACENA` | Kupac -> Vozac | prijemnica (`KolAmbVracena`) |
| `ROBA_IZDATA_KUPCU` | Kupac -> Vozac | izlaz kupcima |
| `POVRAT_FIRMI` | Stanica -> Firma | revers `IZDATO_OM` |
| `PRIJEM_OD_FIRME` | Firma -> Stanica | revers `PRIJEM_OD_OM` |

Izvestaj tada nikad ne pita „znaci li `OM-Izlaz-Koop` revers ili izdate prazne",
nego pita `VrstaKretanja = PRAZNA_IZDATA`.

### 6.4 Nalozi: `Firma` nije granica sistema

| Nalog | Znacenje | Saldo ima fizicko znacenje |
|---|---|---|
| `Kooperant` · `Stanica` · `Kupac` · `Vozac` | stvarni drzaoci gajbica | da |
| `Firma` | **centralni magacin** — stvaran drzalac | **da** |
| `SpoljniSvet` | granica: nabavka, otpis, lom, gubitak | **ne** — to je izvor/ponor |

`Firma` ne sme da znaci istovremeno „centralni magacin", „odnekud su se pojavile
gajbice" i „ovde su nestale polomljene" — tada njen saldo nema jednu semantiku.
Nabavka je `SpoljniSvet -> Firma`, otpis je `bilo koji nalog -> SpoljniSvet`.
Nijedan od ta dva dogadjaja **ne postoji u kodu danas** (provereno); model im
ostavlja mesto bez novog mehanizma.

**Identitet naloga ostaje par `Tip + ID`, ali uz kapiju.** Zasebna tabela naloga
(opaque `NalogID`) bila bi cistija protiv para „`Tip=Vozac`, `ID=KUP-17`", ali
uvodi **peti registar** koji mora da prati cetiri master tabele — a to je tacno
klasa koja je ovaj repo vec ujela (`MRTAV_UNOS` u `vba_hard_census`). Nalog
nezavisan od entiteta nije danasnja potreba. Zato:

> **AMB-10-ODL-1.** Par `Tip + ID` se razresava **iskljucivo** kroz jednu
> fail-closed kapiju u `PrenesiAmbalazu`; nijedan pisac ne proverava tipove sam.
> Vrata za tabelu naloga ostaju otvorena: `OdNalogID`/`NaNalogID` bi je zamenili
> bez promene oblika dogadjaja.

### 6.5 Saldo: jedna formula, bez izuzetaka

```
saldo(nalog) = SUM Kolicina WHERE Na = nalog
             - SUM Kolicina WHERE Od = nalog
```

Nema `entitetTip` grananja, nema inverzije, nema „ko je transporter". Vozacev
saldo ispada sam — jer je vozac **nalog**, ne izuzetak.

> Danasnja inverzija je **fail-open**: citalac koji zaboravi
> `VozacAmbEffectiveSmer` ne dobija gresku nego **pogresan znak**. U ciljnom
> modelu tu gresku nije moguce napraviti, jer inverzije nema.

### 6.6 Storno je kontra-stav, ne zastavica

Storno ne menja postojeci red nego upisuje nov, sa zamenjenim stranama i
`StornoOd` koji pokazuje na original. Time odluka o stampi **prestaje da bude
odluka**:

| Pitanje | Upit |
|---|---|
| sta je dokument tada rekao | dogadjaji tog `DokumentID` **bez** kontra-stavova |
| sta vazi danas | **svi** dogadjaji, ukljucujuci kontra-stavove |

Pisac storna **ne prima** strane, kolicinu ni tip od pozivaoca:

```
StornirajPrenos(originalAmbID)   ' sve ostalo cita IZ ORIGINALA
```

Pozivalac koji sme da posalje svoj iznos sme i da posalje pogresan.

Posledica za registar: `tblAmbalaza` ulazi u `modSchemaGuard.BEZ_STORNA`, ali
**iz drugog razloga** nego stavke tabele — ne zato sto status drzi zaglavlje,
nego zato sto je knjiga nepromenljiva. Taj razlog mora da stoji u registru,
inace sledeci citalac spoji dve razlicite stvari pod istim imenom.

### 6.7 Invarijante — sedam, i nijedna tautoloska

| ID | Tvrdnja |
|---|---|
| `AMB-INV-01` | `Kolicina > 0` |
| `AMB-INV-02` | `OdNalog <> NaNalog` |
| `AMB-INV-03` | oba naloga postoje i dozvoljena su za tu `VrstaKretanja` |
| `AMB-INV-04` | **isti poslovni efekat ne postoji dvaput**: `(DokumentID, VrstaKretanja, TipAmbalaze)` je jedinstven medju nestorniranim dogadjajima |
| `AMB-INV-05` | storno je **tacan inverz**: ista kolicina i tip, zamenjene strane |
| `AMB-INV-06` | jedan original ima **najvise jedan** storno; storno se ne stornira |
| `AMB-INV-07` | nalozima kojima minus nije dozvoljen saldo ne sme pasti ispod nule — **koji su to nalozi je poslovna odluka**, nije pretpostavljeno |

> **Povucena tvrdnja.** Prva verzija je kao glavnu invarijantu nudila „zbir salda
> po svim nalozima je konstantan". **To je tautologija**: svaki prenos po
> konstrukciji daje `-x` i `+x`, pa je globalni zbir nula i kad je dogadjaj
> dupliran i kad su strane pogresne. Ostaje kao sanity check nad **oblikom
> podatka**, ne kao dokaz ispravnosti.

### 6.8 Idempotencija: invarijanta da, nova kolona ne

Review je trazio `OperationID` uz `EventRole`, da retry iste komande ne duplira
fizicki transfer. **Zahtev se prihvata, mehanizam ne** — i to zbog merenja:

**svih devet knjizenja su UNUTAR dokumentove transakcije** koja snapshot-uje
`tblAmbalaza` (`IzdajOtpremnicu_TX:3469`, `IspravkaOtpremnice_TX:3548`,
prijemnica `:6729`, izlaz kupcima `:7021`, otkup — sopstveni komentar
`modOtkup:89` „Ambalaza je u snapshotu zbog pada IZMEDJU dva TrackAmbalaza
poziva"). Red knjige zato **ne moze da prezivi neuspeo upis dokumenta**;
polu-upisano stanje koje bi retry zatekao ne postoji.

Duplikat je zato **dokumentski**, ne knjigovodstveni: da bi se isti transfer
upisao dvaput, mora postojati drugi uspesan dokument — a identitet dokumenta je
vec cuvan (`ClientRecordID` za PWA ingest, registar brojeva i storna za desktop).

Posledica: stabilan identitet poslovnog efekta **vec postoji** i glasi
`(DokumentID, VrstaKretanja, TipAmbalaze)`. To je `AMB-INV-04`. Nova kolona bi
bila drugi identitet iste stvari — a dva identiteta jedne stvari su tacno ono sto
T1 u sekciji 4 prijavljuje kao kvar.

**Cena ovog izbora, izgovorena:** ako se ikad pojavi knjizenje **van** dokumentove
transakcije (npr. mrezni poziv koji sam pravi dogadjaj), ovaj kljuc vise nije
dovoljan i `OperationID` tada ulazi. Do tada bi bio nemerena odbrana.

### 6.9 Sta je review promenio

| Nalaz | Ishod |
|---|---|
| `AMB-INV-01` je tautologija | **prihvaceno — moja greska.** Zamenjeno sa sedam invarijanti (6.7) |
| `ModifiedAt`/`ModifiedBy` nad nepromenljivom knjigom | **prihvaceno — moja greska**, unutrasnja protivrecnost spec-a. Obrisane |
| `Firma` mesa magacin i granicu | **prihvaceno.** Uveden `SpoljniSvet` (6.4) |
| `EventRole` odvojen od tipa dokumenta | **prihvaceno.** `VrstaKretanja`, zatvoren enum (6.3) |
| storno bez unique/exact-inverse kapije | **prihvaceno.** `AMB-INV-05/06` + pisac koji ne prima iznos (6.6) |
| polimorfni `Tip + ID` | **prihvaceno uz izmenu:** fail-closed kapija umesto petog registra, sa obrazlozenjem i otvorenim vratima (6.4) |
| `OperationID` za retry | **zahtev prihvacen, mehanizam odbijen** uz merenje (6.8) |

### 6.10 Redosled — stare strukture se brisu POSLEDNJE

1. **AMB-10a** — ugovor: nalozi, `VrstaKretanja`, sedam invarijanti. **Bez produkcionog cutovera.**
2. **AMB-10b** — nov append-only pisac + svih devet mesta knjizenja + `AMB-INV-04` + sabotaze.
3. **AMB-10c** — saldo i izvestaji na jednu formulu; staro i novo se mere **jedno protiv drugog** dok oba postoje.
4. **AMB-10d** — storno kao tacan inverz; istorijski i tekuci upit.
5. **AMB-10e** — **tek tada** brisanje: stare kolone, `VozacAmbEffectiveSmer`, `ReversID`, `Stornirano`, kes kolone, `Chk_B10` izuzetak.

Prva verzija je brisala u istom rezu u kom uvodi pisca. I bez legacy podataka to
je prevelik blast radius za jedan rez — korak 10e postoji da bi staro i novo
mogli da se mere jedno protiv drugog pre nego sto staro ode.

**AMB-02, AMB-03 i AMB-04 iz sekcije 4 se povlace kao zasebni predlozi** — sve
troje su posledice AMB-10, a ne rezovi za sebe.

### 6.11 Cena — posteno

Najveci redizajn jedne tabele u refaktoru. Dira: cetiri funkcije salda, tri
izvestaja + karticu, `modIntegritet` (`Chk_B10` i susedi), obe storno putanje,
`modBrojevi` (revers numeracija), stampu, i devet mesta knjizenja. Broj redova
pritom **pada** (otkup 4 -> 2). Jedino sto ga cini jeftinim je isto sto vazi za
ceo refaktor: **nema podataka za migraciju.**

## 7) Raniji predlozi (povuceni -- v. 6.7)

1. **AMB-03** — keš sa dokumenata (traži odgovor na pitanje o stornu). Mali,
   zatvoren, 9 čitalaca.
2. **AMB-04** — taksonomija (poslovna odluka o izveštajima).
3. **AMB-02** — identitet u `DokumentID` (najveći; dira revers i izlaz kupcima).

AMB-03 ne zavisi ni od jednog drugog. AMB-02 je lakši **posle** AMB-04, jer tada
tipovi već razdvajaju put dokumenta od puta reversa.
