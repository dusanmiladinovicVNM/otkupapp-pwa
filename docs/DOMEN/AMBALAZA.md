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

## 6) Ciljni model — AMB-10

> Odluka operatera 27.09.2026: **štampa storniranog dokumenta prikazuje ono što je
> važilo pre storna.** I pitanje: da li ideal traži dve noge u svakom događaju.
>
> **Ne traži.** Dve noge nisu cilj nego **simptom** — posledica reda koji ume da
> imenuje samo jednu stranu. Zato danas isti model proizvodi dva oblika: dva reda
> kad su obe strane nosioci salda, jedan red plus **pravilo** kad je druga strana
> vozač. Ideal ne dodaje drugu nogu — **ukida pojam noge.**

### 6.1 Događaj je jedan red koji imenuje obe strane

```
tblAmbalaza  (događaj = PRENOS)

  AmbID              identitet događaja
  Datum
  TipAmbalaze
  Kolicina           UVEK pozitivna -- smera nema, smer je (Od -> Na)

  OdNalogTip         Kooperant | Stanica | Kupac | Vozac | Firma
  OdNalogID
  NaNalogTip         isto
  NaNalogID

  DokumentTIP        povod
  DokumentID         UVEK identitet, nikad broj

  StornoOd           AmbID događaja koji ovaj poništava (kontra-stav); inače prazno
  CreatedAt/By, ModifiedAt/By
```

**Šta nestaje, ne menja se nego nestaje:** `Smer` · `EntitetID`/`EntitetTip` ·
`VozacID` · `ReversID` · `Stornirano` · cela funkcija
`VozacAmbEffectiveSmer` · izuzetak u `Chk_B10` · obe keš kolone
(`tblOtkup.KolAmbIzdata`, `tblPrijemnica.KolAmbVracena`) · druga storno putanja.

### 6.2 Zašto prenos, a ne dve noge

| | dve noge (klasično dvojno) | **prenos (Od → Na)** |
|---|---|---|
| „tačno dve strane" | traži **kapiju** koja proverava da noge postoje i da se slažu — to je `Chk_B10` i danas | **strukturno nemoguće prekršiti**: siroče noge nema jer noge nema |
| uparivanje | traži identitet događaja na svakom redu | red **jeste** događaj |
| broj redova | otkup: 4 reda za 2 događaja | otkup: **2 reda** |
| smer | kolona koja se može pogrešno upisati | **izveden iz para**, ne postoji kao podatak |

Ovo nije nov model nego **dovršen postojeći**: današnji red već *misli* obe
strane — jednu upisuje, drugu podrazumeva. Ideal je prestati podrazumevati.

### 6.3 Saldo: jedna formula, bez izuzetaka

```
saldo(nalog) = SUM Kolicina WHERE Na = nalog
             - SUM Kolicina WHERE Od = nalog
```

Nema `entitetTip` grananja, nema inverzije, nema „ko je transporter". Vozačev
saldo ispada sam — jer je vozač **nalog**, ne izuzetak. `GetVozacAmbSaldo`,
`GetStanicaAmbSaldo` i `GetAmbalazeStanje` postaju jedan poziv sa drugim
argumentom.

> Današnja inverzija je **fail-open**: čitalac koji zaboravi
> `VozacAmbEffectiveSmer` ne dobija grešku nego **pogrešan znak**. U ciljnom
> modelu tu grešku nije moguće napraviti, jer inverzije nema.

### 6.4 Storno je kontra-stav, ne zastavica

Knjiga je **nepromenljiva**: storno ne menja postojeći red nego upisuje nov, sa
zamenjenim stranama i `StornoOd` koji pokazuje na original.

Time odluka o štampi **prestaje da bude odluka**:

| Pitanje | Upit |
|---|---|
| šta je dokument tada rekao | događaji tog `DokumentID` **bez** kontra-stavova |
| šta važi danas | **svi** događaji, uključujući kontra-stavove |

Oba iz istog podatka, bez tumačenja zastavice — a to je tačno ono što je
traženo za štampu storniranog dokumenta.

Posledica za registar: `tblAmbalaza` ulazi u `modSchemaGuard.BEZ_STORNA`, ali
**iz drugog razloga** nego stavke tabele — ne zato što status drži zaglavlje,
nego zato što je knjiga nepromenljiva. Taj razlog mora da stoji u registru,
inače sledeći čitalac spoji dve različite stvari pod istim imenom.

### 6.5 Očuvanje postaje merljiva invarijanta

Prvi put se može tvrditi:

> **AMB-INV-01.** Zbir salda po **svim** nalozima je konstantan. Gajbica ne
> nastaje i ne nestaje — samo menja nalog.

Danas se to **ne može** tvrditi, jer jednonožni događaji legalno „cure": kad OM
vrati ambalažu firmi, upisuje se samo noga stanice i gajbice ispadaju iz knjige.
U ciljnom modelu `Firma` je **nalog**, pa je perimetar zatvoren.

Nabavka novih gajbica i otpis (lom, gubitak) **nisu mereni u kodu** — ne postoje.
Ako se pojave, to su događaji nad nalogom `Firma` i model već ima mesto za njih:
nije potreban nov mehanizam, samo nov nalog ako se poželi razdvajanje.

### 6.6 Identitet: `DokumentID` je uvek identitet

`ReversID` nestaje jer prestaje da bude potreban: revers je **dokument**, dakle
dobija svoj `ReversID` kao identitet i on ide u `DokumentID`, kao što otkup šalje
`OtkupID`. Broj ostaje labela — `ZBR-IDENT-01`, isto pravilo, treći put.

Time i storno ima **jednu** putanju umesto dve: sve se razrešava po
`DokumentID`, a kontra-stav se veže `StornoOd`-om.

### 6.7 Cena — pošteno

Ovo je **najveći redizajn jedne tabele** u celom refaktoru. Dira: četiri funkcije
salda, tri izveštaja + karticu, `modIntegritet` (`Chk_B10` i susedi), obe storno
putanje, `modBrojevi` (revers numeracija), štampu, i devet mesta knjiženja.

Jedino što ga čini jeftinim je isto što važi za ceo refaktor: **nema podataka za
migraciju.** Broj redova pritom **pada** (otkup 4 → 2).

Rez se ne može izvesti u jednom potezu i ne treba ga tako ni planirati:

1. **AMB-10a** — nova šema + pisac: `TrackAmbalaza` postaje `PrenesiAmbalazu(od, na, ...)`, devet mesta knjiženja prelazi na nju.
2. **AMB-10b** — saldo i izveštaji na jednu formulu; `VozacAmbEffectiveSmer` se briše.
3. **AMB-10c** — storno kao kontra-stav; druga putanja se briše; `tblAmbalaza` u `BEZ_STORNA` sa svojim razlogom.
4. **AMB-10d** — keš kolone sa dokumenata (`KolAmbIzdata`, `KolAmbVracena`) i `Chk_B10` izuzetak — oboje **ispada samo po sebi**, nije zaseban rez.

**AMB-02, AMB-03 i AMB-04 iz sekcije 4 se time povlače kao zasebni predlozi** —
sve troje su posledice AMB-10, a ne rezovi za sebe. Izvedeni odvojeno bili bi
zakrpe nad modelom koji se ionako menja.

## 7) Raniji predlozi (povuceni -- v. 6.7)

1. **AMB-03** — keš sa dokumenata (traži odgovor na pitanje o stornu). Mali,
   zatvoren, 9 čitalaca.
2. **AMB-04** — taksonomija (poslovna odluka o izveštajima).
3. **AMB-02** — identitet u `DokumentID` (najveći; dira revers i izlaz kupcima).

AMB-03 ne zavisi ni od jednog drugog. AMB-02 je lakši **posle** AMB-04, jer tada
tipovi već razdvajaju put dokumenta od puta reversa.
