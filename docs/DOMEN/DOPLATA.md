# Doplata kooperantu — knjižno odobrenje po kulturi i klasi

> **Status (03.11.2026):** odluke operatera u §2 su donete, struktura iz §3 je
> **potvrđena**. **F1 (šema, pisac, storno) i F2 (kartica, Saldo OM, PWA) su u
> kodu** — statičke provere zelene, ponašanje **NEVERIFIKOVANO** dok `run_vba` ne
> prođe na Windows + Excel (§7). F2 je napisan pre nego što je F1 pušten na
> Excelu. F3–F4 još nisu počeli. Otvoreno: §8.

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

## 2) Odluke operatera

| ID | Odluka |
|---|---|
| **DOP-ODL-01** | Stopa je **din/kg**. Iznos stavke = `ZaokruziNovac(kg × stopa)`. Bez PDV nadoknade — knjižno odobrenje. |
| **DOP-ODL-02** | Ključ kulture je **ključ cenovnika**: `VrstaVoca + SortaVoca + Klasa`, tačno poklapanje kao `modCenovnik.GetVazecaCena` (trim, bez razlike u veličini slova). Prazna sorta znači **kultura bez sorte**, ne „sve sorte". |
| **DOP-ODL-03** | Obuhvat: **svi aktivni otkupi** — sva otkupna mesta, uključujući hladnjačke blokove — u periodu **od–do po datumu otkupa**. |
| **DOP-ODL-03a** | Period ima i **sat** (odluka 03.11.2026). Dan je `Datum` dokumenta; sat je vreme nastanka po pravilu otkupnog lista (`modOtkup.OtkupVremeNastanka`: `SourceCreatedAt`, pa `CreatedAt`). Sat se gleda **samo kad granica seče dan** (Od posle 00:00 prvog dana, Do pre 23:59 poslednjeg). Obe granice su uključene, na minut. Kad `SourceCreatedAt` postoji a nije vreme (PWA ISO tekst), sat je **nepoznat** — ne uzima se `CreatedAt`, jer je to za PWA red trenutak sinhronizacije — i otkup na graničnom danu obara obračun po imenu (§8, stavka 1). |
| **DOP-ODL-04** | Datum zaduženja = **datum knjiženja** doplate. Na kartici kooperanta: **jedan red po kooperantu** po obračunu. |
| **DOP-ODL-05** | Doplata se **isplaćuje kroz postojeće naloge za banku** i automatsko mapiranje izvoda (faza F4). Gotovinska isplata KO ide na **matičnu stanicu** kooperanta (`tblKooperanti.StanicaID`), i kad KO ima stavke na više stanica. |
| **DOP-ODL-06** | **Ponovljena doplata** nad istim ključem i preklopljenim periodom je **dozvoljena**, uz upozorenje koje imenuje raniji KO i izričitu potvrdu operatera. |
| **DOP-ODL-07** | Knjižena doplata je **snimak**: ne prati kasniji storno ili ispravku otkupa. Promena osnovice je **signal** (integritet / Oporavak, faza F3), ne automatska prepravka. Ispravka = storno KO + nov obračun (obrazac ZBR-KANON-03). |
| **DOP-ODL-08** | Izveštaj **Prosečna cena** se ne menja. |
| **DOP-ODL-09** | Broj KO je **`KO-<n>/<godina>`**, niz po godini datuma knjiženja, preko svih KO (i storniranih — storno ne oslobađa broj, A9). |
| **DOP-ODL-10** | *Predlog uz F1, operater ga može oboriti:* kraj perioda nije u budućnosti; datum knjiženja nije pre poslednjeg dana perioda ni posle današnjeg; svaka stopa mora da pogodi bar jedan aktivan otkup (tipfeler u sorti ne sme da da doplatu nula za ceo ključ); prazan obračun se ne knjiži. |
| **DOP-ODL-11** | Kooperantu se **štampa KO** (odluka 03.11.2026) — faza F3. |

## 3) Struktura dokumenta (potvrđena 03.11.2026)

**Jedno knjižno odobrenje (KO) po kooperantu.** Globalni unos je **obračun**,
koji u jednoj transakciji pravi po jedan KO za svakog kooperanta sa kg > 0.

| Tabela | Kolone (kanon: `schema/schema.json`) |
|---|---|
| `tblDoplata` (zaglavlje = jedan KO) | DoplataID, BrojDokumenta, Datum, KooperantID, ObracunID, PeriodOd, PeriodDo, Napomena, Stornirano, CreatedAt, CreatedBy, ModifiedAt, ModifiedBy |
| `tblDoplataStavke` | DoplataStavkaID, DoplataID, RedniBroj, StanicaID, VrstaVoca, SortaVoca, Klasa, Kolicina, Cena, Iznos, CreatedAt, CreatedBy, ModifiedAt, ModifiedBy |

- **Kardinalnost:** 1 obračun : N KO (jedan po kooperantu) ; 1 KO : M stavki
  (stanica × vrsta × sorta × klasa). Kooperant koji je predavao na dva otkupna
  mesta ima **jedan** KO sa stavkama po stanici; Saldo OM čita stavke svoje stanice.
- **Identitet:** `DoplataID` (`DOP-…`) je kanon; `BrojDokumenta` je labela.
  `BrojDokumenta` je pod ugovorom formata `text`, kao u svim tabelama koje ga imaju.
  Slovni prefiks `KO-` drži poziv na broj odvojen od brojeva otkupa:
  `NormalizePozivKey` (`modBankaMapiranje`) zadržava slova i cifre. Proveriti u
  F4 nad svim nizovima koje banka razrešava.
- **`ObracunID`** (`DOB-…`) je zajednički ključ jednog globalnog unosa (obrazac
  `GeneracijaID`) — za pregled i za storno celog obračuna. Nije zasebna tabela.
- **Stope** ne idu u `tblCenovnik` i nisu treći model cene
  (`.claude/rules/agrohemija-i-cene.md`): žive samo na stavkama dokumenta.
- **Kartica (F2):** red nosi opis sa razlaganjem („Doplata Malina Willamette I
  1.250 kg × 10") i ref-ključ `DOP|<DoplataID>` za „Detalje", kao red otkupa.

## 4) Invarijante i vlasnici

| ID | Pravilo | Vlasnik (kod · test) |
|---|---|---|
| **DOP-01** | Doplata ne nosi robu: ne ulazi ni u jedan kg zbir; cena na stavkama otkupa se ne menja. | `modDoplata` ne piše `tblOtkup*` · 309 („plan ne piše"), 310 („cena na stavkama otkupa ostaje netaknuta") |
| **DOP-02** | `Iznos` stavke = `ZaokruziNovac(Kolicina × Cena)`; `Kolicina` = zbir aktivnih `tblOtkupStavke` tog kooperanta, stanice i ključa u periodu, u trenutku knjiženja (čitač `modOtkup.StavkeOtkupaRedovi`, ne nov). | `modDoplata.DoplataPlan` · 309, 311 |
| **DOP-03** | Storno nije brisanje: `tblDoplata` je u `STORNO_TABELE`, `tblDoplataStavke` u `BEZ_STORNA` (stavka prati zaglavlje). Storniran KO izlazi iz svih zbirova; broj ostaje zauzet. | `modStorno.StornoDoplata_TX` / `StornoObracunDoplate_TX`, čitač `modDoplata.DoplataStavkeRedovi` · 313 |
| **DOP-04** | Pisci: nastanak samo `modDoplata` (`KnjiziDoplatu_TX`, obe tabele); storno zaglavlja `modStorno`, kao kod ostalih dokumenata. | `WRITE_OWNERSHIP.json` (A11), `who_writes.py --check-ownership` |
| **DOP-05** | (F4) Novac vezan za KO se pri stornu KO **odvezuje, ne stornira** — isto pravilo kao kod otkupa (`ODLUKA_NOVAC_PRI_STORNU.md`). | F4 |
| **DOP-06** | Jedan KO po kooperantu po obračunu; brojevi uzastopni u nizu godine. | `modDoplata.KnjiziPlan` · 310 |
| **DOP-07** | Doplata je **zaduženje u korist kooperanta** svuda gde se vidi saldo: na kartici je jedan red po KO na datum knjiženja (i u početnom stanju), a u Saldo OM i PWA izvozu saldo = vrednost + doplata − isplaćeno − agro. U Saldo OM stavka KO pripada stanici na kojoj je roba predata, a period je po datumu knjiženja. | `modDoplata.DoplataIznosPoKooperantu` / `DoplataKOKooperanta`, `modIzvestaj`, `modStammdatenSync.SaldoOMDetailRedovi` · 315, 316, 317 |

**Prikaz i pisac dele jedan plan** (obrazac `BimAutoPlan`): `DoplataPlan` računa
stavke bez upisa, ekran ga prikazuje (F3), `KnjiziDoplatu_TX` ga izvršava.

## 5) Poslovni događaji

| Događaj | Datum / tabela |
|---|---|
| fizički | nema — roba se ne kreće |
| poslovni / pravni | KO nastaje na datum knjiženja (`tblDoplata.Datum`) |
| finansijski | obaveza firme prema kooperantu nastaje istog dana (zaduženje na kartici); isplata je poseban, kasniji događaj u `tblNovac` |

Poslovni i finansijski događaj dele datum zato što je to pravilo knjižnog
odobrenja (DOP-ODL-04), ne zbog lakše implementacije.

## 6) Nizvodno — ko čita činjenicu

| Čitač | Promena | Gde se vidi | Faza |
|---|---|---|---|
| `modIzvestaj.ReportKarticaKooperanta` | **urađeno:** red po KO (`DOP\|<DoplataID>`), zaduženje, ulazi u početno stanje; opis `KarticaDoplataOpis` („Doplata … kg x … din"); rang izvora: doplata 6, novac 7 (KART-HRON-01) | ekran kartice, PDF po šablonu, PWA `getKartica` | F2 |
| `modIzvestaj.ReportSaldoOM` | **urađeno:** kolona **8 „Doplata"** (na kraju, pozicioni čitaoci); saldo (6) je uključuje | ekran, tabelarni PDF, KPI pločica (kolona 6) | F2 |
| `modIzvestaj.ReportSaldoOMZbirni` | **urađeno:** kolona **9 „Doplata"**; ekran sabira saldo po indeksu (OM 7, kupci 6) umesto „pretposlednje kolone" | ekran, tabelarni PDF | F2 |
| `modStammdatenSync.SaldoOMDetailRedovi` (seam izdvojen iz `ExportSaldoOMDetail`) | **urađeno:** kolona **10 „Doplata"**; saldo je uključuje; PWA kartica kooperanta u menadžmentu prikazuje doplatu | PWA saldo | F2 |
| otvorene obaveze (`modNovac.GetOpenOtkupi` ili njegov par za KO), `modBankaExportPregled`, `ValidateNalogSaldo`, `BimAutoPlan` + resolver po broju KO | KO kao otvorena obaveza; `tblNovac` dobija kolonu `DoplataID` **na kraju** | nalozi za banku, specifikacija isplata (PDF), uvoz izvoda | F4 |

**Ne menja se:** `ReportIsplata*`, `ReportKarticaRobaRekap`, `ReportProsecnaCena`,
otkupni list, lanac dokumenata. Ekran kartice ne treba izmenu: red bez otkupa
(`DOP|`) već dobija detalj iz opisa i poruku „štampa nedostupna" (štampa KO je F3).

**Raspored promenjen u F2, testovi prilagođeni:** formula u testu Saldo OM
(„saldo = vrednost + doplata − isplaćeno − agro") i ref-kolona zbirnog salda u
dva testa (8 → 9). Sidro sabotaže `kart-hron-izvor-rang` prati novi rang novca (7).

**Nije pročitano (pre F2/F3):** unutrašnjost `ReportSaldoOMZbirni` i
`ReportProsecnaCena`, registracija novog tipa dokumenta u storno toku i ekranu
Storno, PWA `pregled.js`.

## 7) Faze

1. **F1 — šema, pisac, storno, testovi. U KODU (03.11.2026).**
   - `schema/schema.json` (+ `modSchema` regenerisan); `modConfig` (`TBL_DOPLATA*`,
     `COL_DOP*`, `DOP_BROJ_PREFIKS`); `modSchemaGuard` (storno registar);
     `WRITE_OWNERSHIP.json`; `WHO_WRITES.md` regenerisan.
   - `modDoplata`: `DoplataPlan`, `DoplataPreklapanja`, `KnjiziDoplatu_TX`,
     `DoplataStavkeRedovi`, `AktivneDoplateObracuna`, `BrojKO`, `PoslednjiBrojKO`.
   - `modStorno`: `StornoDoplata_TX`, `StornoDoplata`, `StornoObracunDoplate_TX`.
   - `modOtkup`: `OtkupVremeNastanka` (pravilo sata, sada jedno za štampu i
     doplatu — `modPrint.FillOtkupSablon` ga zove, ponašanje štampe isto);
     `RequireValidOtkupClass` je `Public`.
   - `modTest` 309–314 (`T_Doplata_*`), sabotaže `dop-*` u `tools/sabotaza.py`.
   - Nove tabele pravi `EnsureAllTables` pri startu (`modMain.StartApp`).
2. **F2 — čitaoci. U KODU (03.11.2026).** Kartica, Saldo OM (pojedinačno i
   zbirno), PWA izvoz (§6); `modTest` 315–317, sabotaže `dop-kartica-*`,
   `dop-saldo-om-*`, `dop-zbirni-*`, `dop-pwa-*`.
3. **F3 — ekran i papir.** Nov ekran u ljusci `frmOtkupUI` (runtime kontrole, bez
   novih `WithEvents`): period sa satom, tabela stopa, pregled plana sa
   „stope bez otkupa", potvrda ponavljanja, „Knjiži"; štampa KO (DOP-ODL-11);
   vidljivost na ekranu Storno; signal promene osnovice (DOP-ODL-07) u Oporavku.
4. **F4 — isplata.** Nalozi za banku, automatsko mapiranje izvoda,
   `tblNovac.DoplataID`, gotovina na matičnoj stanici (DOP-ODL-05).

**Spuštanje je serijsko:** F1 i F4 diraju centralne registre (`schema.json`,
`modSchemaGuard`, `WRITE_OWNERSHIP.json`, brojevi testova u `modTest`). Na dan
F1 otvoren je plan S8 paleta (PR #456, samo docs); grana koja posle doda testove
od 309 nadalje sudara se sa ovom.

## 8) Otvoreno

1. **⚠ PLATFORM UNKNOWN — sat PWA otkupa.** PWA piše `CreatedAtClient` kao
   `new Date().toISOString()` (UTC, npr. `2026-11-02T07:00:00.000Z`), a uvoz ga
   prenosi u `SourceCreatedAt` kao tekst (`modMasterSync`, `GS_CREATED_AT`). Ako
   VBA `IsDate` taj oblik ne prihvata — što se očekuje, ali **nije izmereno** —
   štampa otkupnog lista za PWA otkup danas pokazuje vreme **sinhronizacije**
   (`CreatedAt`), a doplata takav otkup na graničnom danu odbija (DOP-ODL-03a).
   Test 311 to meri; sonda nad pravom sveskom je u izveštaju uz F1. Ako se
   potvrdi, odluka je operaterova: parsirati ISO + UTC → lokalno vreme u
   `OtkupVremeNastanka` (popravlja i štampu), ili ostaviti granice na celim danima.
2. DOP-ODL-10 je predlog — potvrditi ili oboriti.
