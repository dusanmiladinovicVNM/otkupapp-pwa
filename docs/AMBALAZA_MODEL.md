# Kretanje ambalaže (gajbica) u sistemu

**Svrha:** detaljno objašnjenje kako se ambalaža (gajbice/paleti) kreće kroz OtkupApp —
kako se knjiži, kako se računaju saldi po entitetu i po vozaču, i zašto su neki
dokumenti jednostrani a neki dvojni upisi.

**Companion:** `ARCHITECTURE_REFERENCE.md` (sekcija o ambalaža ledgeru / saldo pravilima).
Kanonski izvor koda: `modAmbalaza` (ledger), `modDokumenta` / `modOtkup` / `frmDokumenta` /
`modMasterSync` (upis), `modIzvestaj` (čitanje/izveštaji), `modStorno` (storno).

---

## 1. Osnovni princip — Smer je ENTITETSKI-relativan

Svaki red u `tblAmbalaza` opisuje **jedno kretanje gajbica iz ugla jednog entiteta**:

- `Smer = "Ulaz"`  → gajbice **ulaze** tom entitetu (on ih dobija / drži).
- `Smer = "Izlaz"` → gajbice **izlaze** iz tog entiteta (on ih predaje).

Entitet je određen poljima `EntitetID` + `EntitetTip` (`Kooperant` / `Stanica` (OM) /
`Kupac`). **Smer NIJE relativan na firmu** — uvek se čita iz ugla entiteta na tom redu.

### Pravila po tipu entiteta

| Entitet | `Ulaz` (dobija gajbice) | `Izlaz` (predaje gajbice) |
|---|---|---|
| **Kooperant** | dobije **prazne** gajbice (od firme/OM) | preda **pune** na OM (otkup) |
| **OM / Stanica** | primi **prazne** od firme; primi **pune** od kooperanta (otkup) | da **pune** vozaču (otpremnica); da **prazne** kooperantu (izdavanje) |
| **Kupac** (hladnjača) | primi **pune** od zbirne (prijemnica) | da **prazne** vozaču (vraćene / izlaz-kupci) |

> Napomena: za **broj gajbica** je svejedno da li su pune ili prazne — gajbica je gajbica.
> „Pune/prazne" je samo fizička priča; ledger broji `Ulaz`/`Izlaz` po komadu.

---

## 2. Ledger: `tblAmbalaza`

Append-only tabela. Relevantne kolone:

| Kolona | Značenje |
|---|---|
| `AmbID` | PK (`AMB-…`) |
| `Datum` | datum kretanja |
| `TipAmbalaze` | npr. `4/1`, `12/1` |
| `Kolicina` | broj gajbica |
| `Smer` | `Ulaz` / `Izlaz` (entitetski-relativno) |
| `EntitetID` / `EntitetTip` | entitet kretanja (`KOOP-…`/`ST-…`/`KUP-…` + `Kooperant`/`Stanica`/`Kupac`) |
| `VozacID` | vozač uključen u kretanje (prazno kad ga nema) |
| `DokumentID` / `DokumentTip` | izvorni dokument (`OTK-…`/`OTP-…`/`PRJ-…` + tip) |
| `Stornirano` | oznaka storna |

Sav upis ide kroz **jednu** proceduru: `modAmbalaza.TrackAmbalaza(datum, tipAmb, kolicina,
smer, entitetID, entitetTip, [vozacID], [dokumentID], [dokumentTip])`. Poziva se iz većih
poslovnih TX wrapper-a koji već snapshot-uju `tblAmbalaza` (zato nema `TrackAmbalaza_TX`).
`TrackAmbalaza` fail-fast validira ulaz i radi `AppendRow`.

---

## 3. Mapa knjiženja po dokumentu

Svako od ovih kretanja je **entitetski-relativno** (vidi §1).

| Dokument | Procedura | Knjiži | Vozač | `DokumentTip` |
|---|---|---|---|---|
| **Otpremnica** | `modDokumenta.SaveOtpremnica` | `Stanica` **Izlaz** | da | `Otpremnica` |
| **Prijemnica** — 1. txt (pune) | `modDokumenta.SavePrijemnica` | `Kupac` **Ulaz** | da | `Prijemnica` |
| **Prijemnica** — 2. txt (zamena/vraćene) | `modDokumenta.SavePrijemnica` | `Kupac` **Izlaz** | da | `Prijemnica` |
| **Izlaz Kupci** | `modDokumenta.SaveKupciIzlaz_TX` | `Kupac` **Izlaz** | da | `Kupci-Otpremnica` |
| **Otkup** (desktop) | `modOtkup.SaveOtkup` | `Kooperant` **Izlaz** + `Stanica` **Ulaz** | — | `Otkup` |
| **Otkup** (PWA sync) | `modMasterSync.ImportRowToTblOtkup` | `Kooperant` **Izlaz** + `Stanica` **Ulaz** | — | `Otkup` |
| **OM izdaje kooperantu** | `frmDokumenta.SaveOMUlaz_TX` (toggle „Izdato koop.") | `Kooperant` **Ulaz** + `Stanica` **Izlaz** | — | `OM-Izlaz-Koop` |
| **OM izdaje kooperantu (uz otkup)** | `modOtkup.SaveOtkup` (polje „Izdata ambalaza" u `frmOtkup`) | `Kooperant` **Ulaz** + `Stanica` **Izlaz** | — | `OM-Izlaz-Koop` |
| **OM prima od kooperanta (povrat prazne)** | `frmDokumenta.SaveOMUlaz_TX` (toggle „Prijem koop.") | `Kooperant` **Izlaz** + `Stanica` **Ulaz** | — | `OM-Ulaz-Koop` |
| **OM-Ulaz** (prijem na OM) | `frmDokumenta.SaveOMUlaz_TX` (default) | `Stanica` **Ulaz** | da | `OMUlaz` |

**Zbirna** ne dira ambalažu (nema upisa).

> **Dvoklasni otkup (Klasa I + Klasa II).** Otkup sa obe klase su **dva `tblOtkup` reda**
> (isti `BrDok`); **svaki red knjiži svoju ambalažu** — `kolAmb` (Klasa I) i `kolAmbII`
> (Klasa II) — kroz isti dvojni upis (`Kooperant Izlaz` + `Stanica Ulaz`). Kod unosa
> **samo Klase II** (bez Klase I) i otkupna i **izdata** ambalaža se knjiže na red Klase II
> (jedini koji postoji). Isto važi za dokumenta (otpremnica/zbirna/prijemnica): Klasa II
> nosi svoje gajbe kroz ceo lanac i paletizaciju. Storno celog dokumenta hvata sve redove
> preko `BrDok` (vidi §9).

---

## 4. Jednostran vs dvojni upis — i zašto

Pitanje: zašto neki dokumenti knjiže **dva** reda, a neki **jedan**?

- **Dvojni upis** se koristi kada kretanje menja saldo **dva realna entiteta** od kojih
  se **nijedan ne može izvesti** iz reda onog drugog:
  - **Otkup**: gajbice idu **kooperant → OM** (pune). `Kooperant Izlaz` (kooperant se
    razdužuje) **+** `Stanica Ulaz` (OM se zadužuje).
  - **OM-izdavanje**: gajbice idu **OM → kooperant** (prazne). `Kooperant Ulaz` (dobija
    prazne) **+** `Stanica Izlaz` (OM se razdužuje).
  - **OM-prijem (povrat prazne)**: gajbice idu **kooperant → OM** (prazne, bez otkupa).
    `Kooperant Izlaz` **+** `Stanica Ulaz` — isti smer kao otkup, ali zaseban `DokumentTip`
    (`OM-Ulaz-Koop`) jer nije nabavka (ne ulazi u otkupne izveštaje/izuzeća).
  - Svi tokovi su između **dva realna entiteta** (kooperant ↔ OM); vozač nije strana.

- **Jednostran upis** se koristi kada je druga strana kretanja **vozač** — a vozač se
  **izvodi pri čitanju** (vidi §5), pa nije potreban drugi red:
  - Otpremnica, Prijemnica, Izlaz-Kupci, OM-Ulaz — svaki knjiži samo svoj entitet
    (`Stanica`/`Kupac`), a vozačka noga se računa naknadno.

> Pravilo: ako je „druga strana" realan entitet sa svojim saldom → dvojni upis.
> Ako je „druga strana" vozač (transporter) → jednostrano + izveden vozač.

---

## 5. Vozač = izvedeni inverzni protivpartner

Vozač (transporter) **nema svoj red** u ledgeru — njegova ambalaža se računa iz redova
koji nose njegov `VozacID`, tako što se **smer invertuje**: što entitetu **uđe** (`Ulaz`),
iz vozača **izlazi** (`Izlaz`), i obrnuto. Vozač je „pokretni magacin": otpremnica ga
**puni** (utovar), prijemnica ga **prazni** (istovar).

Pravilo je u `modAmbalaza.VozacAmbEffectiveSmer(smer, entitetTip)`:

| `EntitetTip` reda | Vozačev smer |
|---|---|
| `Stanica` ili `Kupac` (transport) | **invertovan** (`Izlaz`↔`Ulaz`) |
| `Kooperant` | nepromenjen (i ionako izuzet, vidi §6) |

Izvedeni vozač po dokumentu:

| Dokument (entitet/smer) | Vozač |
|---|---|
| Otpremnica (`Stanica Izlaz`) | **Ulaz** (puni se — utovar na OM) |
| Prijemnica pune (`Kupac Ulaz`) | **Izlaz** (prazni se — istovar kupcu) |
| Prijemnica vraćene (`Kupac Izlaz`) | **Ulaz** (pokupio prazne od kupca) |
| Izlaz-Kupci (`Kupac Izlaz`) | **Ulaz** (pokupio prazne od kupca) |
| OM-Ulaz (`Stanica Ulaz`) | **Izlaz** (vratio prazne na OM) |

Tako **kompletna ruta otpremnica → prijemnica daje vozaču saldo 0**; otvorena otpremnica
(još nije predato) → **pozitivan** saldo = gajbice još kod vozača.

---

## 6. Otkup je izuzet iz vozačkog salda

Otkup (`DokumentTip = "Otkup"`) je **nabavka** (kooperant → OM); **vozač nije strana** tog
kretanja. Iste gajbice se kasnije broje na **otpremnici** (transportna noga), pa bi
brojanje otkupa duplo teretilo vozača. Zato se otkup **izuzima** iz vozačkog salda na
**dva** mesta:

- `modIzvestaj.ReportAmbalaza` (vozač grana): filter `DokumentTip <> "Otkup"`.
- `modAmbalaza.GetVozacAmbSaldo`: preskače redove sa `DokumentTip = "Otkup"`.

> Auto-hladnjača istorijski forsira mirror-vozača (`VozacID == StanicaID`) na svaki
> hladnjača-otkup; izuzimanje znači da to ne kvari vozačev saldo. `tblOtkup.VozacID`
> (grupisanje u zbirnu / sledljivost) ostaje netaknut — samo se **saldo** ne računa.

---

## 7. Životni ciklus gajbice (primeri)

### 7.1 Petlja OM ↔ kooperant (prazne)
| Korak | Kooperant | OM (Stanica) |
|---|---|---|
| OM izdaje prazne kooperantu | **+N** (`Ulaz`) | **−N** (`Izlaz`) |
| Kooperant vrati pune (otkup) | **−N** (`Izlaz`) | **+N** (`Ulaz`) |
| **Neto** | **0** | **0** |

### 7.2 Transportna ruta (OM → kupac)
| Korak | OM | Kupac | Vozač (izveden) |
|---|---|---|---|
| Otpremnica (OM → vozač) | **−N** (`Izlaz`) | — | **+N** (`Ulaz`) |
| Prijemnica (vozač → kupac) | — | **+N** (`Ulaz`) | **−N** (`Izlaz`) |
| **Neto** | −N | +N | **0** |

### 7.3 Vraćanje prazne ambalaže
| Korak | Kupac | OM | Vozač (izveden) |
|---|---|---|---|
| Kupac vraća prazne vozaču | **−N** (`Izlaz`) | — | **+N** (`Ulaz`) |
| Vozač vraća prazne na OM (OM-Ulaz) | — | **+N** (`Ulaz`) | **−N** (`Izlaz`) |
| **Neto** | −N | +N | **0** |

---

## 8. Saldi i izveštaji

- **Entitetski saldo** (`modAmbalaza.GetAmbalazeStanje(entitetID, entitetTip)`): sabira
  redove tog entiteta, `Ulaz = +Kolicina`, `Izlaz = −Kolicina`. Koristi **sirov** `Smer`.
- **Vozačev saldo / izveštaj** (`modAmbalaza.GetVozacAmbSaldo`, `modIzvestaj.ReportAmbalaza`
  za `Vozac`): filtrira po `VozacID`, **izuzme `Otkup`** (§6), i svaki red provuče kroz
  `VozacAmbEffectiveSmer` (§5). Kompletna ruta → saldo 0.
- **Entitetski izveštaji** (`OM` / `Kupac` / `Kooperant`) koriste sirov `Smer`
  (`isVozac = False`).
- **Početno stanje pre bloka** (`modAmbalaza.GetKooperantAmbOpening(koopID, tipAmb, blockOtkupIDs)`):
  entitetski saldo kooperanta (sirov `Smer`) nad redovima upisanim **pre prvog reda datog
  bloka** (granica = najmanji red-indeks gde `DokumentID ∈ blockOtkupIDs`, append-only redosled).
  Koristi ga `modPrint` za red „Saldo ambalaze" na otkupnom listu: `saldo = početno + izdato −
  primljeno` (ispravno i na ponovnoj štampi starijeg bloka — kasniji blokovi se ne uračunavaju).
- Samo **ne-stornirani** redovi ulaze u aktivne saldo helpere.

---

## 9. Storno

`modStorno.StornoAmbalazaByDokument(dokumentID, dokumentTip)` markira **sve** redove tog
dokumenta kao stornirane → **automatski hvata obe noge dvojnog upisa** (otkup i
OM-izdavanje), jer obe noge dele isti `DokumentID` + `DokumentTip`.

**Storno celog otkupnog dokumenta:** `modStorno.StornoOtkupByBrDok_TX(brDok)` stornira **sve
`tblOtkup` redove** istog dokumenta (Klasa I + Klasa II, dele `BrDok`) u jednoj transakciji —
za svaki red poziva `StornoOtkup`, koji reversuje i njegove ambalaža noge (otkupnu **i**
izdatu `OM-Izlaz-Koop`, jer dele `DokumentID = otkupID`). Koristi se iz `frmDokumenta` i
panela (`modOtkupBlok.StornoSelectedBlok`, fallback na `OtkupID`), pa storno dvoklasnog
dokumenta više ne ostavlja drugu klasu nestorniranu.

> Storno-UI: `OM-Izlaz-Koop` (revers izdavanje) i `OM-Ulaz-Koop` (revers povrat) imaju
> putanju u storno comboboxu („Revers izdavanje koop." / „Revers povrat koop.") —
> `modStorno.StornoOMKoopByBrDok_TX(brDok, dokumentTip)` markira **obe noge** po broju
> dokumenta (broj je obavezan; unos bez broja nema jedinstven ključ). Novac unet uz isti
> broj stornira se zasebno („Novac"). Preostala praznina: plain `OMUlaz` (prijem na OM od
> vozača) i dalje nije u storno comboboxu. Napomena: storno OM-koop reversa se još ne
> prikazuje u „Pregled storniranih" (`GetStorniraniByTip` je po prodajnim tabelama).
>
> Izuzetak: `OM-Izlaz-Koop` knjižen **uz otkup** (polje „Izdata ambalaza" u `frmOtkup`)
> deli `DokumentID = otkupID`, pa ga `StornoOtkup` automatski stornira (dodatni
> `StornoAmbalazaByDokument otkupID, DOK_TIP_OM_IZLAZ_KOOP`).

---

## 10. Migracija / istorijski podaci

`tblAmbalaza` je append-only — promena pravila knjiženja važi **samo za nove redove**
(posle re-importa relevantnih modula). Postojeći redovi zadržavaju staru konvenciju:

- **Prijemnica** uneta pre entity-relativne ispravke ima staru orijentaciju (`Kupac Izlaz`
  za pune umesto `Ulaz`).
- **Otkup** unet pre dvojnog upisa nema `Stanica Ulaz` nogu.

Za uskladiti: **re-seed** test podataka, ili jednokratna migracija (flip/dodavanje
nedostajućih redova). Pokrenuti **tačno jednom** i tek **posle** re-importa koda.

---

## 11. Reference (kod)

| Oblast | Gde |
|---|---|
| Ledger / upis | `modAmbalaza.TrackAmbalaza`, `GetAmbalazeStanje`, `GetVozacAmbSaldo`, `VozacAmbEffectiveSmer`, `GetKooperantAmbOpening` |
| Otpremnica / Prijemnica / Izlaz-Kupci | `modDokumenta` (`SaveOtpremnica`, `SavePrijemnica`, `SaveKupciIzlaz_TX`) |
| Otkup | `modOtkup.SaveOtkup` (desktop), `modMasterSync.ImportRowToTblOtkup` (PWA) |
| OM-Ulaz / OM-izdavanje / OM-prijem-koop | `frmDokumenta.SaveOMUlaz_TX` + runtime toggle-i `tglIzdKoop` (izdato) i `tglPrijemKoop` (prijem/povrat); smer = parametar `koopSmer` |
| Broj reversa (auto) | `modBrojevi.SuggestNextBroj(KIND_REV, stanicaID, datum)` → `x/ddmmyy[-N]`; poštuje toggle `AUTO_BROJ_DOKUMENTA` (`IsAutoBrojDokumenta`); sopstveni dnevni niz po stanici (scan `tblAmbalaza`, OM-koop tokovi) |
| Vozač/entitet izveštaji | `modIzvestaj.ReportAmbalaza` (+ `ReportAmbalazePojedinacni`/`Zbirni`) |
| Storno | `modStorno.StornoAmbalazaByDokument`; standalone revers: `StornoOMKoopByBrDok_TX(brDok, dokumentTip)`; pregled: `modDokumenta.GetStorniraniRevers` |
| Konstante tipova | `modConfig` (`DOK_TIP_OTKUP`, `DOK_TIP_OTPREMNICA`, `DOK_TIP_PRIJEMNICA`, `DOK_TIP_IZLAZ_KUPCI`, `DOK_TIP_OM_ULAZ`, `DOK_TIP_OM_IZLAZ_KOOP`, `DOK_TIP_OM_ULAZ_KOOP`) |
