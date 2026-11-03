# Doplata kooperantu — knjižno odobrenje po kulturi i klasi

> **Status (03.11.2026):** odluke operatera u §2 su donete. Struktura dokumenta u
> §3 je **PREDLOG koji čeka potvrdu**. Kod ne postoji; do potvrde §3 i stavki iz
> §8 nema produkcionog koda (pre-flight: `DOMAIN` / `CARDINALITY` su još `GAP`).

## 1) Šta je doplata

Naknadno odobrenje kooperantu. Firma **globalno** zada stopu u **din/kg** po
ključu cenovnika (vrsta + sorta + klasa) za period. Sistem sabere kilograme
aktivnih otkupa svakog kooperanta po tom ključu i knjiži **jedan iznos u korist
kooperanta**.

Vodi se kao **knjižno odobrenje**: iznos koji se obračuna je iznos koji se
isplaćuje, **bez PDV nadoknade**.

Doplata **ne menja**:

- robu (ne nosi kilograme ni u jedan robni zbir);
- cenu na stavkama otkupa (`tblOtkupStavke.Cena`);
- otkupni list (pravni dokument je već odštampan; ponovna štampa mora dati isto);
- hladnjački lanac (cena stavke prijemnice je i dalje cena stavke otkupa, HLD-LANAC-01).

Odbačene alternative:

| Alternativa | Zašto ne |
|---|---|
| Podići `Cena` na stavkama otkupa | menja već odštampan otkupni list i hladnjačku prijemnicu; nije „jedan unos" |
| Red u `tblNovac` | `Tip` u `tblNovac` nosi **kanal plaćanja** (`modConfig.bas`, „Novac Tipovi"); doplata je obaveza, ne novac — svaki čitač uplata/isplata bi morao da je isključuje |

## 2) Odluke operatera (03.11.2026)

| ID | Odluka |
|---|---|
| **DOP-ODL-01** | Stopa je **din/kg**. Iznos stavke = `ZaokruziNovac(kg × stopa)`. Bez PDV nadoknade — knjižno odobrenje. |
| **DOP-ODL-02** | Ključ kulture je **ključ cenovnika**: `VrstaVoca + SortaVoca + Klasa`, tačno poklapanje kao `modCenovnik.GetVazecaCena` (trim, bez razlike u veličini slova). Prazna sorta znači **kultura bez sorte**, ne „sve sorte". |
| **DOP-ODL-03** | Obuhvat: **svi aktivni otkupi** — sva otkupna mesta, uključujući hladnjačke blokove. Period po datumu otkupa (od–do) je **tumačenje koje treba potvrditi** (§8). |
| **DOP-ODL-04** | Datum zaduženja = **datum knjiženja** doplate. Na kartici kooperanta: **jedan red po kooperantu** po obračunu. |
| **DOP-ODL-05** | Doplata se **isplaćuje kroz postojeće naloge za banku** i automatsko mapiranje izvoda (faza F4). |
| **DOP-ODL-06** | **Ponovljena doplata** nad istim ključem i preklopljenim periodom je **dozvoljena**, uz upozorenje koje imenuje raniji obračun i izričitu potvrdu operatera. |
| **DOP-ODL-07** | Knjižena doplata je **snimak**: ne prati kasniji storno ili ispravku otkupa. Promena osnovice je **signal** (integritet / Oporavak), ne automatska prepravka. Ispravka = storno odobrenja + nov obračun (obrazac ZBR-KANON-03). |
| **DOP-ODL-08** | Izveštaj **Prosečna cena** se ne menja. |

## 3) Struktura dokumenta — PREDLOG, čeka potvrdu

**Jedno knjižno odobrenje (KO) po kooperantu.** Globalni unos je **obračun**,
koji u jednoj transakciji pravi po jedan KO za svakog kooperanta sa kg > 0.

Zašto po kooperantu, a ne jedan dokument sa svim kooperantima:

- knjižno odobrenje je dokument prema **jednom partneru**;
- jedan red na kartici = jedan dokument, isto kao red otkupa (`OTK|<OtkupID>`);
- isplata: **poziv na broj = broj KO**, isto kao broj bloka kod otkupa — jedan
  kooperant, jedan broj, bez para „dokument + kooperant";
- storno pogrešno obuhvaćenog kooperanta ne obara ostale.

| Tabela | Kolone (nove tabele; postojeće se u F1 ne diraju) |
|---|---|
| `tblDoplata` (zaglavlje = jedan KO) | DoplataID, BrojDokumenta, Datum, KooperantID, ObracunID, PeriodOd, PeriodDo, Napomena, Stornirano, CreatedAt, CreatedBy, ModifiedAt, ModifiedBy |
| `tblDoplataStavke` | DoplataStavkaID, DoplataID, StanicaID, VrstaVoca, SortaVoca, Klasa, Kolicina, Cena, Iznos, CreatedAt, CreatedBy, ModifiedAt, ModifiedBy |

- **Kardinalnost:** 1 obračun : N KO (jedan po kooperantu) ; 1 KO : M stavki
  (stanica × vrsta × sorta × klasa). Kooperant koji je predavao na dva otkupna
  mesta ima **jedan** KO sa stavkama po stanici; Saldo OM čita stavke svoje stanice.
- **Identitet:** `DoplataID` je kanon; `BrojDokumenta` je labela. Niz po godini sa
  **slovnim prefiksom** (npr. `KO-17/2026`): `NormalizePozivKey`
  (`modBankaMapiranje`) zadržava slova i cifre, pa prefiks drži poziv na broj
  odvojen od brojeva otkupa. Proveriti u F4 nad svim nizovima koje banka razrešava.
- **`ObracunID`** je zajednički ključ jednog globalnog unosa (obrazac
  `GeneracijaID`) — za pregled i za storno celog obračuna. Nije zasebna tabela.
- **Stope** ne idu u `tblCenovnik` i nisu treći model cene
  (`.claude/rules/agrohemija-i-cene.md`): žive samo na stavkama dokumenta.
- **Kartica:** red nosi opis sa razlaganjem („Doplata Malina Willamette I
  1.250 kg × 10") i ref-ključ `DOP|<DoplataID>` za „Detalje", kao red otkupa.

## 4) Invarijante (planirane; vlasnik = `modDoplata` + test u `modTest`)

| ID | Pravilo |
|---|---|
| **DOP-01** | Doplata ne nosi robu: ne ulazi ni u jedan kg zbir (rekap robe, otpremnica, zbirna, prosečna cena, PWA roba). |
| **DOP-02** | `Iznos` stavke = `ZaokruziNovac(Kolicina × Cena)`; iznos KO = zbir stavki. `Kolicina` = zbir aktivnih `tblOtkupStavke` tog kooperanta, stanice i ključa u periodu, u trenutku knjiženja (čitač `modOtkup.StavkeOtkupaRedovi`, ne nov). |
| **DOP-03** | Storno nije brisanje: `tblDoplata` u `STORNO_TABELE`, `tblDoplataStavke` u `BEZ_STORNA` (stavka prati zaglavlje). Storniran KO izlazi iz svih salda. |
| **DOP-04** | Jedan pisac: samo `modDoplata` piše obe tabele (`WRITE_OWNERSHIP.json`, A11). |
| **DOP-05** | (F4) Novac vezan za KO se pri stornu KO **odvezuje, ne stornira** — isto pravilo kao kod otkupa (`ODLUKA_NOVAC_PRI_STORNU.md`). |

## 5) Poslovni događaji

| Događaj | Datum / tabela |
|---|---|
| fizički | nema — roba se ne kreće |
| poslovni / pravni | KO nastaje na datum knjiženja (`tblDoplata.Datum`) |
| finansijski | obaveza firme prema kooperantu nastaje istog dana (zaduženje na kartici); isplata je poseban, kasniji događaj u `tblNovac` |

Poslovni i finansijski događaj dele datum zato što je to pravilo knjižnog
odobrenja (DOP-ODL-04), ne zbog lakše implementacije.

## 6) Nizvodno — ko čita činjenicu

| Čitač | Promena | Gde se vidi |
|---|---|---|
| `modIzvestaj.ReportKarticaKooperanta` | nov izvor zaduženja; ulazi i u početno stanje; nov rang izvora u `KarticaIzvorRang` (KART-HRON-01) | ekran kartice, PDF po šablonu, PWA `getKartica` |
| `modIzvestaj.ReportSaldoOM`, `ReportSaldoOMZbirni` | nova kolona „Doplata", ulazi u saldo | tabelarni PDF |
| `modStammdatenSync.OtkupSaldoPoKooperantu` / `ExportSaldoOMDetail` | doplata u vrednost | PWA saldo |
| F4: otvorene obaveze (`modNovac.GetOpenOtkupi` ili njegov par za KO), `modBankaExportPregled`, `ValidateNalogSaldo`, `BimAutoPlan` + resolver po broju KO | KO kao otvorena obaveza; `tblNovac` dobija kolonu `DoplataID` **na kraju** | nalozi za banku, specifikacija isplata (PDF), uvoz izvoda |

**Ne menja se:** `ReportIsplata*`, `ReportKarticaRobaRekap`, `ReportProsecnaCena`,
otkupni list, lanac dokumenata.

**Nije pročitano (pre F2/F3):** unutrašnjost `ReportSaldoOMZbirni` i
`ReportProsecnaCena`, registracija novog tipa dokumenta u storno toku i ekranu
Storno, PWA `pregled.js`.

## 7) Faze

1. **F1 — šema i pisac.** `schema/schema.json` → `gen_schema_module.py`;
   `modSchemaGuard` (storno registar); `WRITE_OWNERSHIP.json`; `modDoplata`
   (obračun, knjiženje u jednoj TX, storno); testovi.
2. **F2 — čitaoci.** Kartica, Saldo OM (pojedinačno i zbirno), PWA izvoz.
3. **F3 — ekran.** Nov ekran u ljusci `frmOtkupUI` (runtime kontrole, bez novih
   `WithEvents`): period, tabela stopa, pregled, „Knjiži"; vidljivost na ekranu Storno.
4. **F4 — isplata.** Nalozi za banku, automatsko mapiranje izvoda, `tblNovac.DoplataID`.

**Spuštanje je serijsko:** F1 i F4 diraju centralne registre (`schema.json`,
`modSchemaGuard`, `WRITE_OWNERSHIP.json`). Pre F1 proveriti otvorene grane koje
menjaju iste fajlove (03.11.2026 otvoren je plan S8 paleta, PR #456 — docs).

## 8) Otvoreno — čeka operatera

1. Potvrda strukture iz §3 (KO po kooperantu + `ObracunID`).
2. Period: da li obračun bira period od–do po datumu otkupa, ili uvek obuhvata sve otkupe do dana knjiženja (DOP-ODL-03).
3. Format broja KO (predlog `KO-<n>/<godina>`).
4. F4: kooperant čiji KO ima stavke na dva otkupna mesta, a isplaćuje se gotovinom na otkupnom mestu — kojoj stanici ide isplata.
5. Da li kooperantu treba odštampan KO. Nije traženo: zahtev je prikaz u postojećim PDF izveštajima.
