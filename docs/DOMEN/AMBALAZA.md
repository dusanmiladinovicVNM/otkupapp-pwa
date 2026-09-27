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

## 5) Redosled, ako se ide na sve

1. **AMB-03** — keš sa dokumenata (traži odgovor na pitanje o stornu). Mali,
   zatvoren, 9 čitalaca.
2. **AMB-04** — taksonomija (poslovna odluka o izveštajima).
3. **AMB-02** — identitet u `DokumentID` (najveći; dira revers i izlaz kupcima).

AMB-03 ne zavisi ni od jednog drugog. AMB-02 je lakši **posle** AMB-04, jer tada
tipovi već razdvajaju put dokumenta od puta reversa.
