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

## 6) Ciljni model — AMB-10 (v3)

> **v1** je predlozila prenos umesto nogu. **v2** je zatvorila prvi dizajn review
> (tautoloska invarijanta, `Modified*` nad nepromenljivom knjigom, `Firma` koja
> meša magacin i granicu, `VrstaKretanja`, storno kapije). **v3** zatvara drugi
> krug: **nijedan realni nalog ne sme zavrsiti sa negativnim saldom**, a tudja
> ambalaza ulazi u opticaj **eksplicitnim dogadjajem**, ne cutanjem.
>
> Sta je koji krug promenio: 6.12.

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
  VrstaKretanja      ZASTO je transfer nastao (ne ko kome -- to nose Od/Na)

  StornoOd           AmbID originala koji ovaj ponistava; inace prazno

  CreatedAt
  CreatedBy
```

**Nema `ModifiedAt`/`ModifiedBy`** — knjiga se ne menja. Pogresan unos se ne
popravlja UPDATE-om nego **stornom i novim dogadjajem**.

**Nestaje, ne menja se:** `Smer` · `EntitetID`/`EntitetTip` · `VozacID` ·
`ReversID` · `Stornirano` · `VozacAmbEffectiveSmer` · izuzetak u `Chk_B10` ·
obe kes kolone · druga storno putanja.

### 6.2 Zasto prenos, a ne dve noge

| | dve noge | **prenos (Od -> Na)** |
|---|---|---|
| „tacno dve strane" | trazi **kapiju** da noge postoje i da se slazu — to je `Chk_B10` i danas | **strukturno nemoguce prekrsiti** |
| broj redova | otkup: 4 reda za 2 dogadjaja | **2 reda** |
| smer | kolona koja se moze pogresno upisati | **izveden iz para** |

### 6.3 `SpoljniSvet` je granica opticaja, ne partner

Ambalaza **nastaje** u knjizi samo kao `SpoljniSvet -> neko`, a **izlazi iz
opticaja** samo kao `neko -> SpoljniSvet`. Nije firma, nije magacin, nije
partner — njegov saldo nema fizicko znacenje.

| Nalog | Saldo ima fizicko znacenje |
|---|---|
| `Kooperant` · `Stanica` · `Kupac` · `Vozac` | **da** — stvarni drzaoci, po entitetu |
| `Firma` | **da** — **jedan jedini nalog** (odluka operatera 28.09.2026); magacini se ne razdvajaju |
| `SpoljniSvet` | **ne** — izvor i ponor |

### 6.4 Negativan saldo ne postoji — deficit se POKRIVA, ne trpi

> **Nijedan realni nalog ne sme posle commit-a imati saldo < 0.**

To ne znaci da kooperant ili kupac ne smeju doneti svoje gajbe. Znaci da taj
dogadjaj mora **eksplicitno da poveca** kolicinu u pracenom opticaju.

**Kooperant donosi svoje** (`saldo 5`, predaje `20`):

```
SpoljniSvet -> K1   15   ULAZ_TUDJE_AMBALAZE
K1 -> Stanica       20   AMBALAZA_UZ_ROBU
------------------------------------------
K1 = 0      Stanica += 20      firma duguje K1: 15
```

**Kupac je isti slucaj u ogledalu** (`saldo 20`, vraca `30` praznih):

```
SpoljniSvet -> KUP3  10   ULAZ_TUDJE_AMBALAZE
KUP3 -> Vozac        30   POVRAT_PRAZNE
------------------------------------------
KUP3 = 0    Vozac += 30      firma duguje KUP3: 10
```

Ovo **nije edge case** nego osnovno pravilo: ista domenska operacija na oba kraja
lanca.

### 6.5 Protokol potvrde deficita — kapija je u PISCU, ne u UI-ju

```
PrenesiAmbalazu(...)                -- sam racuna deficit
    deficit = 0                     -> upisuje
    deficit > 0, bez potvrde        -> AMB_EXTERNAL_CONFIRM_REQUIRED(deficit)
    deficit > 0, potvrda != deficit -> ODBIJA, pita ponovo
    deficit > 0, potvrda == deficit -> dva dogadjaja u ISTOJ TX
```

Pisac **ponovo racuna stanje** pri drugom pozivu i ne veruje staroj potvrdi:
izmedju pitanja i odgovora stanje se moglo promeniti drugim unosom.

> To nije nov obrazac nego **isti koji repo vec drzi**: „Modul unosa proverava
> nad snimkom iz trenutka kad je lista punjena... zato kriticna poslovna kapija
> stoji i u writer-u" (`.claude/rules/testovi.md` §5, kao
> `ApplyAvansToOtkup` / `IsplataBlokProblem`).

Odbijanje je **potpuno**: nema dokumenta, nema ambalaze, nema parcijalnog upisa.

### 6.6 Fizicko stanje i dug vlasniku nisu ista stvar

Knjiga odgovara na „gde su gajbe". Obaveza se **izvodi** iz iste knjige, bez
ijedne mutabilne kolone:

```
Obaveza(partner, tip) = SUM ULAZ_TUDJE_AMBALAZE     (SpoljniSvet -> partner)
                      - SUM VRACANJE_TUDJE_AMBALAZE (firma -> partner)
```

Zato **`VRACANJE_TUDJE_AMBALAZE` mora biti odvojeno od `IZDATA_PRAZNA`**, iako
su fizicki isti potez (`Stanica -> Kooperant`):

| | Fizicki | Obaveza |
|---|---|---|
| `IZDATA_PRAZNA` | koop += N | firma ga **zaduzuje** svojim gajbama |
| `VRACANJE_TUDJE_AMBALAZE` | koop += N | firma **zatvara pozajmicu** |

**To je i dokaz da `VrstaKretanja` mora postojati kao podatak:** dva dogadjaja sa
istim `Od`, `Na` i `Kolicina` nose suprotno poslovno znacenje. Iz para se ne moze
izvesti.

### 6.7 `VrstaKretanja` — domenski pojmovi, ne spisak call-site-ova

| Vrednost | Znacenje |
|---|---|
| `AMBALAZA_UZ_ROBU` | gajbe putuju **sa robom** (otkup, otpremnica, prijemnica, izlaz kupcu) |
| `IZDATA_PRAZNA` | firma zaduzuje partnera praznim gajbama |
| `POVRAT_PRAZNE` | partner vraca prazne gajbe firmi |
| `PRENOS_INTERNO` | izmedju sopstvenih naloga (stanica <-> firma) |
| `ULAZ_TUDJE_AMBALAZE` | partnerove gajbe ulaze u opticaj — **stvara obavezu** |
| `VRACANJE_TUDJE_AMBALAZE` | firma vraca partneru njegove — **gasi obavezu** |
| `NABAVKA` | nove gajbe ulaze u opticaj (firmine) |
| `OTPIS` | lom, gubitak — izlaze iz opticaja |
| `POCETNO_STANJE` | **firmine** gajbe zatecene kod entiteta pri uvodjenju — zaduzuje ga, **ne** stvara obavezu |

Nema `OTKUP_*` ni `PRIJEMNICA_*` (to kaze `DokumentTIP`) ni
`STANICA_KOOPERANT` (to kazu `Od`/`Na`).

> **`POCETNO_STANJE` je dodato na osnovu merenja, ne iz review-a.**
> `GetKooperantAmbOpening` danas **ne cita nikakvo pocetno stanje** nego ga
> sabira **iz same knjige** ([modAmbalaza.bas:358](../../src-vba/modAmbalaza.bas)).
> U novom modelu prvi dan zato pocinje na nuli: svaka izdata gajbica bila bi
> deficit i svaki unos bi trazio potvrdu. Zatecena kolicina mora da udje kao
> dogadjaj.
>
> **ODLUKA OPERATERA (28.09.2026).** `POCETNO_STANJE` se odnosi **iskljucivo na
> ambalazu u vlasnistvu firme** koja je na dan uvodjenja zatecena kod entiteta za
> koji se pocetno stanje radi. Dakle:
>
> ```
> SpoljniSvet -> Entitet    POCETNO_STANJE
>   -> entitet je ZADUZEN firminim gajbama
>   -> obaveza firme prema njemu NE nastaje
> ```
>
> **Posledica koju treba izgovoriti:** partnerove **sopstvene** gajbe zatecene na
> dan uvodjenja **nisu** pocetno stanje. One ulaze kao `ULAZ_TUDJE_AMBALAZE` i
> stvaraju obavezu — isto kao da su donete sutra. Time su dve uloge razdvojene i
> na dan uvodjenja, ne samo u toku rada.
>
> Zato `POCETNO_STANJE` i `ULAZ_TUDJE_AMBALAZE` ostaju dve vrednosti iako su
> fizicki isti prenos: razlikuje ih **vlasnistvo**, a ono odlucuje o obavezi.

### 6.8 Saldo i storno

```
saldo(nalog) = SUM Kolicina WHERE Na = nalog - SUM Kolicina WHERE Od = nalog
```

Nema grananja po tipu, nema inverzije, vozac ispada sam jer je **nalog**.
Danasnja inverzija je fail-open: citalac koji zaboravi `VozacAmbEffectiveSmer`
dobija **pogresan znak**, ne gresku.

Storno je kontra-stav:

```
StornirajPrenos(originalAmbID)     ' i nista vise
```

Pozivalac **ne salje** strane, kolicinu ni tip — pisac ih cita iz originala.
Pozivalac koji sme da posalje svoj iznos sme i da posalje pogresan.

| Pitanje | Upit |
|---|---|
| sta je dokument **tada** rekao | dogadjaji tog `DokumentID` **bez** kontra-stavova |
| sta vazi **danas** | **svi** dogadjaji |

`tblAmbalaza` ulazi u `modSchemaGuard.BEZ_STORNA`, ali **iz drugog razloga** nego
stavke tabele — ne zato sto status drzi zaglavlje, nego zato sto je knjiga
nepromenljiva. Razlog mora da stoji u registru.

### 6.9 Invarijante

| ID | Tvrdnja |
|---|---|
| `AMB-INV-01` | `Kolicina > 0` |
| `AMB-INV-02` | `OdNalog <> NaNalog` |
| `AMB-INV-03` | oba naloga jednoznacno razresena kroz **jednu** kapiju `ResolveAmbNalog(Tip, ID)` — `Tip=Vozac, ID=KUP-123` pada pre upisa |
| `AMB-INV-04` | za originalni dogadjaj je **`(DokumentTIP, DokumentID, VrstaKretanja, TipAmbalaze)`** jedinstven; isti identitet + isti sadrzaj = idempotentno, isti identitet + drugi sadrzaj = **HARD CONFLICT** |
| `AMB-INV-05` | storno je **tacan inverz**; pozivalac ne salje vrednosti |
| `AMB-INV-06` | jedan original ima **najvise jedan** storno; storno se ne stornira |
| `AMB-INV-07` | **nijedan realni nalog nema saldo < 0** posle commit-a; deficit je dozvoljen samo ako je u **istoj TX** pokriven prenosom iz `SpoljniSvet`. Potvrdjeno 28.09.2026: vazi za SVE realne naloge, bez izuzetka |
| `AMB-INV-08` | **nijedan upis u knjigu ne nastaje van vlasnistva transakcije izvornog dokumenta** |
| `AMB-INV-09` | **`Obaveza(partner, tip) >= 0`** — firma ne moze partneru vratiti vise tudje ambalaze nego sto je od njega uzela |

Zbir svih salda ostaje **sanity check nad oblikom**, ne dokaz ispravnosti: svaki
prenos po konstrukciji daje `-x` i `+x`, pa je nula i kad je dogadjaj dupliran.

#### `AMB-INV-04`: zasto i `DokumentTIP`, i sta je „jedan dokument"

Prva verzija kljuca nije nosila `DokumentTIP` i time se **precutno oslanjala** na
to da su svi `DokumentID`-evi u AGRIx-u u jednom globalnom prostoru imena. Kolona
vec stoji na redu knjige, pa oslanjanje nema cenu koju bi platilo — `DokumentTIP`
ulazi u kljuc.

Drugo pitanje istog kljuca: **sme li jedan dokument da proizvede dva dogadjaja sa
istom `VrstaKretanja` i `TipAmbalaze`?** Za danasnjih devet tokova ne sme i ne
dešava se. Ali `tblAmbalazaDokument` je genericki, pa bi jedan `POCETNO_STANJE`
dokument nad **dva entiteta** to odmah prekrsio. Zato:

> **AMB-10-ODL-3.** Jedan `tblAmbalazaDokument` pokriva **tacno jednog
> protivpartnera** — kao sto revers vec danas pokriva jednog kooperanta. Pocetno
> stanje za pet entiteta je pet dokumenata, ne jedan sa pet redova.

Alternativa bi bila prosiriti kljuc nalozima, ali tada on prestaje da bude
identitet **poslovnog efekta** i postaje identitet reda — a idempotencija se meri
po efektu.

#### `AMB-INV-09`: obaveza je storno-svesna, i ima dno

Formula iz 6.6 (`SUM ULAZ - SUM VRACANJE`) **nije bila tacna nad append-only
knjigom**: storno `ULAZ_TUDJE_AMBALAZE` upisuje kontra-stav, fizicki saldo se
anulira, a obaveza bi ostala da visi. Zato se obaveza racuna preko **doprinosa**,
a ne preko dve sume:

```
DoprinosObavezi(dogadjaj):
    ULAZ_TUDJE_AMBALAZE       -> +Kolicina
    VRACANJE_TUDJE_AMBALAZE   -> -Kolicina
    storno (StornoOd != "")   -> MINUS doprinos ORIGINALA
    ostalo                    ->  0

Obaveza(partner, tip) = SUM DoprinosObavezi(svi dogadjaji partnera i tipa)
```

Tako kontra-stav gasi obavezu isto kao sto gasi fizicko stanje — jednim pravilom,
bez posebnog slucaja.

**Donja granica nije kozmetika.** Bez `AMB-INV-09` model moze da proizvede
matematicki ispravan **fizicki** ledger i istovremeno **nemogucu** knjigu obaveza:
firma duguje 12, vrati 20 kao `VRACANJE_TUDJE_AMBALAZE`, saldo prolazi, obaveza
postane `-8`.

> **ODLUKA OPERATERA (28.09.2026): deli se.** Kad se vraca vise nego sto firma
> duguje, pisac **sam** cepa potez na dva dogadjaja:
>
> ```
> obaveza = 12, vraca se 20
>   Stanica -> K1   12   VRACANJE_TUDJE_AMBALAZE   (gasi obavezu na 0)
>   Stanica -> K1    8   IZDATA_PRAZNA             (novo zaduzenje partnera)
> ```
>
> Podelu radi **pisac**, ne operater i ne UI: ista kolicina, isti partner, isti
> trenutak — a granica je `Obaveza(partner, tip)` koju samo pisac cita pouzdano.
> UI sme da je **prikaze** unapred („od 20 izdatih: 12 povrat, 8 novo zaduzenje"),
> ali racun koji vazi je onaj iz pisca, u trenutku upisa — isti razlog kao kod
> potvrde deficita (6.5).
>
> Dva dogadjaja, ne jedan sa dva znacenja: `AMB-INV-04` ih razlikuje po
> `VrstaKretanja`, a `AMB-INV-09` posle oba i dalje vazi.

#### `AMB-INV-08`: kapija mora da dokaze CELU tvrdnju

Na `AMB-INV-08` stoji odluka da `OperationID` ne ulazi u model (6.10), pa ne sme
da zivi kao komentar — invarijanta bez kapije je zelja (`pre-flight` §3).

Prva skica kapije trazila je samo `AddTableSnapshot TBL_AMBALAZA` uz svako pozivno
mesto. **To je slabije od same invarijante:** dokazuje da se **knjiga** moze
rollback-ovati, ali ne i da knjiga i njen izvorni dokument imaju **istog vlasnika
transakcije**. Kod koji commit-uje dokument, pa u **novoj** transakciji snapshot-uje
samo knjigu, prosao bi zelen a invarijantu prekrsio.

Kapija zato mora da dokaze svih pet:

1. postoji obuhvatna poslovna transakcija;
2. `TBL_AMBALAZA` je u njenom snapshotu;
3. nastanak ili izmena **izvornog dokumenta** pripada **istoj** transakciji;
4. za `tblAmbalazaDokument` — i ta tabela je u istoj transakciji;
5. za zatecene putanje (OTK/OTP/PRJ) checker nosi **registar vlasnistva**, isto
   kao `who_writes` za upis redova.


### 6.10 `OperationID`: zahtev prihvacen, mehanizam odbijen

Izmereno: **svih devet knjizenja su unutar dokumentove transakcije** koja
snapshot-uje `tblAmbalaza` (`IzdajOtpremnicu_TX:3469`, `IspravkaOtpremnice_TX:3548`,
prijemnica `:6729`, izlaz kupcima `:7021`, otkup — `modOtkup:89`). Red knjige ne
moze preziveti neuspeo upis dokumenta, pa stanje koje bi retry zatekao ne postoji.

Stabilan identitet efekta zato **vec postoji**: `(DokumentID, VrstaKretanja,
TipAmbalaze)` = `AMB-INV-04`. Nova kolona bila bi drugi identitet iste stvari.

**Uslov pod kojim ovo pada, zapisan unapred:** async red, offline retry,
background posao ili spoljni API koji knjizi ambalazu **sam** — tada `AMB-INV-08`
vise ne vazi i `OperationID` ulazi. Nov takav put mora ponovo otvoriti ovo
pitanje, i zato kapija iz 6.9 postoji.

### 6.11 `KupciIzlaz` nije dokument — to je **revers od kupca**, plus uplata

> **Ispravka operatera (28.09.2026):** „zar kupci izlaz nije u stvari klasican
> revers od kupca ka firmi?" — **jeste**, i merenje to potvrdjuje bez ostatka.
> Ovo obara **i moju preporuku (a) i zahtev review-a za `KupciIzlazID`**.

`SaveKupciIzlaz_TX` ([modDokumenta:6993](../../src-vba/modDokumenta.bas)):

| Mereno | Nalaz |
|---|---|
| kolicina robe | **ne postoji** — u potpisu nema nijednog kg |
| sta prima | `kolAmb` (gajbe) i/ili `novac`; `If kolAmb <= 0 And novac <= 0 Then` pada |
| ko ga zove | **`modNovacUnos` (F6, unos novca)** — „`UplataValidiraj` / `UplataUpisi` F6, `SaveKupciIzlaz_TX` (samo novac)" |
| sta radi sa gajbama | `TrackAmbalaza ... "Izlaz", kupacID, "Kupac", vozacID` — **prazne izlaze od kupca** |
| sta radi sa novcem | `SaveNovac(... fakturaID:=fakturaID ...)` — uplata po **fakturi** |

Dakle to nije dokument nego **ekran poravnanja sa kupcem**, koji u jednoj
transakciji upisuje **dve vec postojece stvari**:

```
1. revers            kupac -> firma      (prazne gajbe)   -> tblAmbalazaDokument
2. uplata            po fakturi          (novac)          -> tblNovac, fakturaID
```

> **AMB-10-ODL-4.** `KupciIzlaz` **ne dobija svoj dokument**. Njegova ambalazna
> polovina je **revers**, i identitet joj je `AmbDokID`; novcana polovina zadrzava
> postojecu vezu (`fakturaID`), gde `brojDok` ostaje **labela**.
>
> Korak **`AMB-10-KI` time nestaje** — ulazi u `10-DOK`.

Ovo je **manji** rez od prethodne odluke, i tacniji: umesto novog dokumenta,
priznaje se da dokument vec postoji i da mu je samo falila tabela.


### 6.11a Razlaganje: dve nezavisne operacije, ne jedan komponovan potez

> **Predlog operatera (28.09.2026):** razdvojiti `KupciIzlaz` na deo za ambalazu i
> deo za novac, pa **revers uciniti jedinstvenim** i **uplate/isplate
> jedinstvenim**.

Prihvaceno **do kraja**: ne samo da `SaveKupciIzlaz_TX` prestaje da bude pisac,
nego prestaje da bude **pojam**. Nema ni dokumenta, ni zajednickog broja, ni
zajednicke transakcije.

| | Ambalaza | Novac |
|---|---|---|
| ulaz | revers kupca | F6, unos novca |
| pisac | `PrenesiAmbalazu` | `SaveNovac_TX` |
| dokument | `tblAmbalazaDokument`, `Vrsta = REVERS` | `tblNovac` + `fakturaID` |
| broj | **svoj** broj reversa | **svoj** `brojDok` |
| transakcija | **svoja** | **svoja** |
| lifecycle / storno | **svoj** | **svoj** |

#### Zasto nije komponovanje u jednoj transakciji (ispravka v5)

v5 je predlagala da ekran zadrzi **jednu** transakciju nad oba pisca. To je
pogresno, i razlog nije stil nego **lifecycle**:

> storno uplate **ne vraca gajbe**, a storno reversa **ne vraca novac**.

Dve cinjenice ciji su zivotni ciklusi nezavisni ne smeju da dele transakciju, jer
zajednicka transakcija podrazumeva **zajednicku sudbinu** koju one nemaju. Sprega
koja se ovde ostavi vraca se kasnije kao pitanje „sta znaci stornirati pola
dokumenta".

**Atomicnost tu nije invarijanta nego udobnost UI-ja.** Ako gajbe legnu a novac ne
(ili obrnuto), **nijedna invarijanta nije prekrsena**: knjiga je tacna, kasa je
tacna, a operater unese polovinu koja fali. Nema stanja koje bi trebalo popraviti.

#### Zajednicki broj je bio GRESKA, i to je pravilo, ne artefakt

Prvo merenje je pokazalo da `SaveKupciIzlaz_TX` prima **jedan** `brojDok` i
prosledjuje ga **obema** stranama ([modDokumenta:7026](../../src-vba/modDokumenta.bas)
i `:7043`). Zapisao sam to kao „artefakt presiroke helper funkcije", jer se iz koda
poslovno pravilo ne moze izvesti.

> **Presuda operatera (28.09.2026):** to je **postojalo, ali je bilo greska.**
> U praksi **nema dokumenta koji pokriva i gotovinsko kretanje novca i ambalazu.**

Time ovo prestaje da bude opis zatecenog stanja i postaje pravilo:

> **AMB-10-ODL-5.** Nijedan dokument nije istovremeno **ambalazni i novcani**. Revers nosi **samo** ambalazu; kasa nosi **samo** novac. Svaki ima svoj broj, svoj identitet, svoju transakciju i svoj storno.
>
> Puna formulacija, sa klasama dokumenata i onim sto **nije** prekrsaj, stoji odmah ispod.

#### Isto pravilo odmah nalazi drugi prekrsaj — ogledalni

Pravilo je opste, pa sam ga primenio na sva mesta koja diraju obe tabele:

| Pisac | Ambalaza | Novac | Deli broj? |
|---|---|---|---|
| `SaveKupciIzlaz_TX` (kupac -> firma) | da | `SaveNovac` | **da — prekrsaj** |
| **`SaveOMUlaz_TX`** (revers ka kooperantu/firmi) | da | `SaveNovac` | **da — prekrsaj** |
| `SaveOtkup*` | da | da | **ne** — v. nize |

**`SaveOMUlaz_TX` je isti kvar u ogledalu:** potpis nosi `novac` i `tipNovca`, ima
istu kapiju `If kolAmb <= 0 And novac <= 0`, i isti `brojDok` deli izmedju
ambalaznih nogu i reda u kasi ([modDokumenta:8555](../../src-vba/modDokumenta.bas)).
Dakle nije rec o izuzetku kod kupca nego o **jednom kvaru na dva mesta**, i
razlaganje se radi na **oba**: revers gubi svoju novcanu polovinu isto kao
`KupciIzlaz`.

#### Prekrsaj je mesanje KLASA, ne dodirivanje dve tabele

Potvrdjeno od operatera (28.09.2026): **otkup nije prekrsaj, i isto vazi za
prijemnicu** — i ona je **robni** dokument koji na sebi nosi ambalazno kretanje.
Pravilo se zato ne izgovara kao „dokument ne sme da dira dve tabele" nego kao
klasifikacija:

| Klasa | Dokumenti | Nosi | Identitet traga |
|---|---|---|---|
| **robni** | otkup, otpremnica, **zbirna**, prijemnica | **robu**; ambalaza (i novac kod otkupa) su njegove **posledice** | `OtkupID` · `OtpremnicaID` · `ZbirnaID` · `PrijemnicaID` |
| **ambalazni** | revers, pocetno stanje, nabavka, otpis | **samo** ambalazu | `AmbDokID` |
| **novcani** | uplata / isplata (kasa) | **samo** novac | `tblNovac` + `fakturaID` |

> **AMB-10-ODL-5 (puno).** Dokument pripada **tacno jednoj** klasi. Robni dokument
> **sme** da proizvede i ambalazni i novcani trag — jer je to **jedan** poslovni
> dogadjaj sa dve posledice, i njemu je izvorni dokument. Prekrsaj je kad
> **ambalazni** dokument nosi novac (ili obrnuto): tada **dva nezavisna dogadjaja**
> dele jedan broj.

Mereno po klasama:

| Pisac | Ambalaza | Novac | Klasa | Ishod |
|---|---|---|---|---|
| `SaveOtkup*` | da | da (`TBL_NOVAC` u snapshotu, `modOtkup:93`) | robni | **u redu** |
| `IzdajOtpremnicu_TX` | da | ne (nema `TBL_NOVAC`, `:3416`) | robni | **u redu** |
| `CreateZbirna_TX` | **ne** | ne | robni | **u redu** — v. nize |
| `SavePrijemnicaMulti_TX` | da | ne (faktura da, kasa ne, `:6565`) | robni | **u redu** |
| `SaveOMUlaz_TX` | da | **da** | ambalazni | **prekrsaj** |
| `SaveKupciIzlaz_TX` | da | **da** | ambalazni | **prekrsaj** |

Razlika koja sve odlucuje: kod robnog dokumenta **jedan dogadjaj ima dva traga**;
kod reversa su **dva dogadjaja delila jedan broj**.


#### Dobitak

- **jedan revers za sve parove**: stanica <-> kooperant, stanica <-> firma,
  kupac -> firma. Cetiri smera plus poseban slucaj na drugom mestu postaju jedan
  mehanizam;
- **jedna kasa**: danas unos **samo novca** na F6 ide kroz funkciju imenovanu po
  ambalazi (`SaveKupciIzlaz_TX`) — to je i bio prvi znak da su dve stvari slepljene;
- `DOK_TIP_IZLAZ_KUPCI` **nestaje** kao tip dokumenta; ti redovi postaju obican revers;
- svaki mehanizam ima svoj `_TX` omotac, po zatecenom obrascu repoa
  (`SaveNovac`/`SaveNovac_TX`, `SavePrijemnica`/`SavePrijemnica_TX`).


### 6.11b Zbirna — jedini dokument lanca koji NE knjizi ambalazu

Dodato zbog celovitosti (predlog operatera, 28.09.2026): zbirna je jedina karika
lanca koju analiza nije izricito svrstala.

**Nalaz 1 — zbirna ne knjizi nijedno kretanje ambalaze, i to je tacno.** Medju
devet izmerenih mesta knjizenja (§3) nema nijednog za zbirnu. To nije propust nego
posledica domena:

```
otpremnica:  Stanica -> Vozac      gajbe odlaze sa stanice
zbirna:      (nista)               grupisanje otpremnica za jedan prevoz
prijemnica:  Vozac  -> Kupac       gajbe stizu kupcu
```

U trenutku zbirne gajbe su **vec kod vozaca** i tu ostaju dok ih prijemnica ne
preda kupcu. Zbirna je **grupisanje**, ne kretanje. Zato je ona **robni dokument
bez ambalazne posledice** — sto klasifikacija iz `AMB-10-ODL-5` dozvoljava: robni
dokument **sme** da ima ambalazni trag, ne **mora**.

> Ovo je ujedno provera same klasifikacije: da je pravilo glasilo „svaki robni
> dokument knjizi ambalazu", zbirna bi ga oborila. Ne obara ga.

**Nalaz 2 — zaglavlje zbirne nosi `UkupnoAmbalaze`, a pisac ga ne puni.** Kanon
`tblZbirna` i dalje ima `UkupnoKolicina`, `UkupnoAmbalaze` i `Klasa`, ali
`CreateZbirna_TX` ih **namerno ostavlja prazne** od PR3 — kolicina po klasi zivi u
`tblZbirnaStavke`, a citaoci idu kroz kanonski helper
([modDokumenta:2409](../../src-vba/modDokumenta.bas)).

To su **mrtvi slotovi**, ista klasa kao linijske kolone zaglavlja otpremnice koje
je obrisao S3-ostatak. Ali:

> **Ne pripada AMB-10 i ne usisava se u njega.** `UkupnoAmbalaze` nije kes
> **kretanja** ambalaze (knjiga o njemu ne zna nista) nego mrtav zbir **stavki
> zbirne**. To je ostatak **S4** — zbirna cutover — i tu se i zatvara, zajedno sa
> `UkupnoKolicina` i `Klasa`, jednim `ObrisiKolonuAko` blokom.
>
> Zapisano je ovde samo da se ne izgubi: nalaz nadjen u ambalaznoj analizi, a dug
> u tudjem rezu.

### 6.12 Sta je koji krug promenio

| Krug | Nalaz | Ishod |
|---|---|---|
| 1 | `AMB-INV-01` tautologija · `Modified*` nad append-only knjigom | **moje greske**, ispravljene |
| 1 | `Firma` mesa magacin i granicu · `VrstaKretanja` · storno kapije · polimorfni `Tip+ID` | prihvaceno (kapija umesto petog registra) |
| 1 | `OperationID` | zahtev prihvacen, mehanizam odbijen uz merenje |
| 2 | **negativan saldo ne postoji**; deficit se pokriva iz `SpoljniSvet` | prihvaceno — `AMB-INV-07`, 6.4 |
| 2 | obaveza za tudju ambalazu je **izveden** read-model | prihvaceno — 6.6 |
| 2 | `VRACANJE_TUDJE_AMBALAZE` odvojeno od `IZDATA_PRAZNA` | prihvaceno; ujedno **dokaz zasto `VrstaKretanja` mora postojati** |
| 2 | enum normalizovan na domenske pojmove | prihvaceno — moj spisak je bio izveden iz call-site-ova |
| 2 | potvrda deficita mora biti backend-safe | prihvaceno — 6.5, isti obrazac kao `IsplataBlokProblem` |
| 2 | `AMB-INV-08` kao tvrd arhitektonski uslov | prihvaceno **i pojacano**: dobija staticku kapiju |
| 2 | `KupciIzlazID` | prihvaceno **uz zaostravanje merenjem**: to znaci **dokument**, 6.11 |
| 4 | podela pri prekomernom vracanju (12 + 8) · razlaganje `KupciIzlaz`-a na **jedan revers** i **jednu kasu** | prihvaceno — 6.9, 6.11a |
| 5 | zajednicki composer/TX je i dalje sprega | **prihvaceno** — dve nezavisne operacije, svaka sa svojom TX (6.11a) |
| 5 | premisa „nikad isti broj" | **operater je presudio: postojalo je, ali je bilo GRESKA.** Postaje `AMB-10-ODL-5`, a merenje po njemu odmah nalazi **drugi, ogledalni prekrsaj** — `SaveOMUlaz_TX` |
| 3 | `Firma` = **jedan** nalog; `POCETNO_STANJE` dobija dokument | prihvaceno; odgovor je izvukao nalaz da **ni revers nema tabelu** — 6.12a |
| — | `POCETNO_STANJE` | **dodato iz merenja**, nije trazeno: `GetKooperantAmbOpening` pocetno stanje izvodi iz knjige, pa bi prvi dan bio zid potvrda |

### 6.12a Ambalazni dokument — jedan, za sve sto svoj nema

> **Odluka operatera (28.09.2026): `POCETNO_STANJE` dobija dokument.**
> Pitanje je bilo uze, ali odgovor je izvukao nalaz koji ga cini sirim.

> **Ispravka operatera:** revers ide i **od stanice ka kooperantu**, ne samo firma <-> stanica. Merenje se slaze: `SaveOMUlaz_TX` ima **cetiri** smera (`IZDAVANJE`, `PRIJEM`, `IZDATO_OM`, `PRIJEM_OD_OM`). Uz 6.11 se dodaje i peti par — **kupac -> firma**. Revers je dakle **partner-genericki** dokument predaje ambalaze, ne interni.

**Mereno:** `ReversID` postoji **samo kao kolona na `tblAmbalaza`** — tabele
reversa **nema nigde u kanonu**. Revers dakle ima identitet i broj, ali **nema
red**. To je **isti plutajuci identitet** koji je u 6.11 odbijen za `KupciIzlaz`,
samo stariji. Zato `DokumentID` kod reversa danas i nosi **broj**: nema cemu da
pokaze.

Iz toga sledi jedan odgovor za sve:

> **AMB-10-ODL-2.** Svi dogadjaji bez sopstvenog poslovnog dokumenta dobijaju
> **jedan zajednicki dokument** — `tblAmbalazaDokument` — cija `Vrsta` kaze sta
> je. Njegov ID ide u `tblAmbalaza.DokumentID`. **`ReversID` time nestaje:** ne
> brise se nego **postaje identitet dokumenta**.

Skica (finalizuje je `10a`):

```
tblAmbalazaDokument
  AmbDokID         identitet -> ide u tblAmbalaza.DokumentID
  Vrsta            REVERS | POCETNO_STANJE | NABAVKA | OTPIS
  BrojDokumenta    labela (modBrojevi; revers zadrzava KIND_REV)
  Datum
  StanicaID        kontekst nastanka
  Napomena
  Stornirano       dokument je dokument -- STORNO_REGISTAR ga ocekuje
  CreatedAt/By, ModifiedAt/By
```

Dokument **nije** knjiga: on sme da nosi `Stornirano` i `Modified*`, jer je
zaglavlje. Njegov storno upisuje **kontra-stavove** u knjigu; knjiga ostaje
append-only. Dve razlicite stvari, dva razlicita ugovora — i to mora da stoji
napisano, jer su u istoj temi.

**Prosirenje koje sam ja izveo, i izgovaram ga da bi se moglo oboriti:**
odluka je trazena za `POCETNO_STANJE`, ali `NABAVKA` i `OTPIS` su **ista klasa** —
dogadjaji bez izvornog dokumenta — pa bi im izuzetak od `AMB-INV-08` bio jedini
alternativni odgovor. Izuzetak u invarijanti je tacno ono sto ovaj rez uklanja iz
`Chk_B10`, pa ih vodim istim putem. Ako je za neku od njih poslovni odgovor
drugaciji, menja se **spisak `Vrsta`**, ne model.

Time u celom domenu ambalaze **nema nijednog dogadjaja bez identiteta dokumenta**:

| Dogadjaj | Dokument |
|---|---|
| otkup, otpremnica, prijemnica | vec postoji |
| ~~`KupciIzlaz`~~ | **nije dokument** — revers + uplata (6.11) |
| **revers** (stanica <-> kooperant, stanica <-> firma, **kupac -> firma**), pocetno stanje, nabavka, otpis | **`tblAmbalazaDokument`** |

`AMB-INV-04` i `AMB-INV-08` tek time vaze **bez ijednog imenovanog izuzetka**.

### 6.13 Redosled — stare strukture se brisu POSLEDNJE

1. **AMB-10a** — ugovor: nalozi + resolver, `SpoljniSvet`, `VrstaKretanja`, `INV-01..09`, protokol potvrde deficita, storno-svesna formula obaveze i njena donja granica. **Bez produkcionog cutovera.**
2. **AMB-10-DOK** — `tblAmbalazaDokument` (revers, pocetno stanje, nabavka, otpis); `ReversID` postaje njegov identitet. Preduslov za `AMB-INV-08` bez izuzetaka.
3. **AMB-10b** — nov append-only pisac + svih devet mesta + pokrivanje deficita + kapije identiteta + sabotaze. **Ukljucuje razlaganje OBA slozena pisca** (`SaveOMUlaz_TX`, `SaveKupciIzlaz_TX`): ambalaza ostaje, novcana polovina odlazi u kasu.
4. **AMB-10c** — saldo, vozac, kooperant, stanica, kupac, ukupno u opticaju, pozajmljeno od partnera; staro i novo se mere **jedno protiv drugog**.
5. **AMB-10d** — storno kao tacan inverz; istorijski i tekuci upit.
6. **AMB-10e** — **tek tada** brisanje starog modela.

**Cetiri dokaza pre `10b`:** zatvoren `VrstaKretanja` enum · tacan protokol
potvrde deficita · test da pozajmljena ambalaza moze **uci -> kretati se -> biti
vracena vlasniku** bez ijednog negativnog realnog salda · **i bez obaveze koja
ikad padne ispod nule ili preživi sopstveni storno**.

### 6.14 Sta model sada ume da odgovori

```
Gde su gajbe?                        saldo po nalogu
Koliko ih je ukupno u opticaju?      zbir realnih naloga
Ko je za njih zaduzen?               saldo partnera
Koliko firma duguje partnerima?      izvedeno iz VrstaKretanja
```

Sve iz jednog append-only modela, bez ijedne rucne kes kolone.

### 6.15 Cena — posteno

Najveci redizajn jedne tabele u refaktoru: cetiri funkcije salda, tri izvestaja +
kartica, `modIntegritet`, obe storno putanje, `modBrojevi`, stampa, devet mesta
knjizenja — plus nov dokument za `KupciIzlaz`. Broj redova pritom **pada**
(otkup 4 -> 2). Jeftinim ga cini samo to sto **nema podataka za migraciju**.

## 7) Raniji predlozi (povuceni -- v. 6.7)

1. **AMB-03** — keš sa dokumenata (traži odgovor na pitanje o stornu). Mali,
   zatvoren, 9 čitalaca.
2. **AMB-04** — taksonomija (poslovna odluka o izveštajima).
3. **AMB-02** — identitet u `DokumentID` (najveći; dira revers i izlaz kupcima).

AMB-03 ne zavisi ni od jednog drugog. AMB-02 je lakši **posle** AMB-04, jer tada
tipovi već razdvajaju put dokumenta od puta reversa.
