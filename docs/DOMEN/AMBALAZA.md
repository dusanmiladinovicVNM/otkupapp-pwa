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
| `Kooperant` · `Stanica` · `Kupac` · `Vozac` · `Firma` | **da** — stvarni drzaoci |
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
| `POCETNO_STANJE` | zateceno stanje pri uvodjenju — **v. merenje nize** |

Nema `OTKUP_*` ni `PRIJEMNICA_*` (to kaze `DokumentTIP`) ni
`STANICA_KOOPERANT` (to kazu `Od`/`Na`).

> **`POCETNO_STANJE` je dodato na osnovu merenja, ne iz review-a.**
> `GetKooperantAmbOpening` danas **ne cita nikakvo pocetno stanje** nego ga
> sabira **iz same knjige** ([modAmbalaza.bas:358](../../src-vba/modAmbalaza.bas)).
> U novom modelu prvi dan zato pocinje na nuli: svaka izdata gajbica bila bi
> deficit i svaki unos bi trazio potvrdu. Zatecena kolicina mora da udje kao
> dogadjaj.
>
> **Otvoreno pitanje za operatera:** gajbe koje partner drzi na dan uvodjenja —
> jesu li **firmine** (partner je zaduzen, obaveza ne nastaje) ili **njegove**
> (obaveza nastaje)? Zato `POCETNO_STANJE` **nije** isto sto i
> `ULAZ_TUDJE_AMBALAZE`, iako je fizicki isti prenos.

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
| `AMB-INV-04` | za originalni dogadjaj `(DokumentID, VrstaKretanja, TipAmbalaze)` je **jedinstven**; isti identitet + isti sadrzaj = idempotentno, isti identitet + drugi sadrzaj = **HARD CONFLICT** |
| `AMB-INV-05` | storno je **tacan inverz**; pozivalac ne salje vrednosti |
| `AMB-INV-06` | jedan original ima **najvise jedan** storno; storno se ne stornira |
| `AMB-INV-07` | **nijedan realni nalog nema saldo < 0** posle commit-a; deficit je dozvoljen samo ako je u **istoj TX** pokriven prenosom iz `SpoljniSvet` |
| `AMB-INV-08` | **nijedan upis u knjigu ne nastaje van vlasnistva transakcije izvornog dokumenta** |

Zbir svih salda ostaje **sanity check nad oblikom**, ne dokaz ispravnosti: svaki
prenos po konstrukciji daje `-x` i `+x`, pa je nula i kad je dogadjaj dupliran.

> **`AMB-INV-08` dobija kapiju, ne obecanje.** Na njemu stoji odluka da
> `OperationID` ne ulazi u model (6.10), pa ne sme da zivi kao komentar. Predlog:
> staticka provera u `tools/` koja za **svako** pozivno mesto `PrenesiAmbalazu`
> trazi `AddTableSnapshot TBL_AMBALAZA` u obuhvatnoj transakciji — isti oblik
> kao `vba_selfupdate_gates`. Invarijanta bez kapije je zelja
> (`pre-flight` §3).

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

### 6.11 `KupciIzlaz`: „stabilan ID" znaci DOKUMENT, ne plutajuci ID

Review trazi `KupciIzlazID`, jer danas `DokumentID` nosi **broj**
([modDokumenta:7026](../../src-vba/modDokumenta.bas)) pa `AMB-INV-04` nije tacna
za sve dogadjaje. Nalaz je tacan; merenje ga zaostrava:

`SaveKupciIzlaz_TX` **ne pise nijednu dokument tabelu** — samo `tblAmbalaza`,
`tblNovac` i `tblFakture`. Dakle ne postoji red kome bi `KupciIzlazID` pripadao.

Zato „dati mu stabilan ID" ima samo dva oblika:

| | Sta je to zapravo |
|---|---|
| **(a)** `KupciIzlaz` dobija **svoj dokument** (red, broj, storno, identitet) | pravi dokument — i konzistentno sa ostatkom: svaki poslovni dogadjaj koji menja robu **i** novac **i** fakturu kod nas ima dokument |
| (b) ID koji zivi samo na redovima knjige | to je **`OperationID` pod drugim imenom** — tacno ono sto je u 6.10 odbijeno |

> **Preporuka: (a).** `KupciIzlaz` je jedina poslovna operacija u sistemu koja
> menja robu, novac i fakturu **bez sopstvenog dokumenta**. To je rupa nezavisna
> od ambalaze; AMB-10 je samo otkriva.
>
> **Blokira `10b`** za taj jedan put: dok dokumenta nema, `AMB-INV-04` bi morala
> da se izgovori sa imenovanim izuzetkom — a izuzetak u invarijanti je ono sto
> smo upravo uklonili iz `Chk_B10`.

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
| — | `POCETNO_STANJE` | **dodato iz merenja**, nije trazeno: `GetKooperantAmbOpening` pocetno stanje izvodi iz knjige, pa bi prvi dan bio zid potvrda |

### 6.13 Redosled — stare strukture se brisu POSLEDNJE

1. **AMB-10a** — ugovor: nalozi + resolver, `SpoljniSvet`, `VrstaKretanja`, `INV-01..08`, protokol potvrde deficita, semantika obaveze, identitet `KupciIzlaz`. **Bez produkcionog cutovera.**
2. **AMB-10b** — nov append-only pisac + svih devet mesta + pokrivanje deficita + kapije identiteta + sabotaze.
3. **AMB-10c** — saldo, vozac, kooperant, stanica, kupac, ukupno u opticaju, pozajmljeno od partnera; staro i novo se mere **jedno protiv drugog**.
4. **AMB-10d** — storno kao tacan inverz; istorijski i tekuci upit.
5. **AMB-10e** — **tek tada** brisanje starog modela.

**Cetiri dokaza pre `10b`:** zatvoren `VrstaKretanja` enum · stabilan identitet
`KupciIzlaz` · tacan protokol potvrde deficita · test da pozajmljena ambalaza moze
**uci -> kretati se -> biti vracena vlasniku** bez ijednog negativnog realnog salda.

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
