# SOFTEK — feature teardown i poređenje sa AgriX-om

**Status:** Drugi feature-level teardown konkurenta u repou
**Datum analize:** 2026-08-02
**Izvor:** `SOFTEK_uputstvp_otkup_poljoproizvoda.pdf` (34 strane), korisničko uputstvo za modul „Otkup poljoprivrednih proizvoda"
**Analiza AgriX strane:** čitanje koda u `src-vba/`, `src/`

---

## 1. Izvor i njegova ograničenja

`PRODUCT DOCUMENTATION EVIDENCE`: korisničko uputstvo za **SOFTEK** modul otkupa
(www.softek.rs). Metapodaci PDF-a: Microsoft Word 2010, **kreirano 2017-06-06**.

`LIMITATION`:

- uputstvo pokriva **samo modul otkupa** (34 strane, pretežno screenshotovi). SOFTEK
  po `competitor_references.csv` ima 13 referenci za „softver za otkup poljoprivrednih
  proizvoda", ali ovo uputstvo ne opisuje ceo proizvod;
- iz sadržaja se vidi da modul stoji **na knjigovodstvenoj platformi** (KEP knjiga,
  kontni nalog za knjiženje, poslovne knjige) — ta platforma nije dokumentovana ovde;
- dokument je iz 2017; aktuelna verzija može biti šira. Sve „nema" tvrdnje znače
  **„nije opisano u ovom uputstvu"**.

---

## 2. Zašto je ovaj dokaz važniji od AGROSOFT-a

`INFERENCE`: **SOFTEK je prvi dokumentovani direktni konkurent AgriX-a.** Za razliku od
AGROSOFT-a (žito, silos, laboratorija, zadruge), SOFTEK-ovo uputstvo radi tačno posao
koji AgriX radi:

- radni primer kroz celo uputstvo je **`MALINA VILAMET I KLASA`** i ambalaža
  **`GAJBICA MALINE`** (str. 11);
- obračun neto količine je **bruto − (broj gajbica × težina gajbice)** (str. 11);
- ambalaža se vodi kroz **revers** i ima `ZADUŽENO / RAZDUŽENO / STANJE` na otkupnom
  listu (str. 12, 16);
- uvod objašnjava **PDV nadoknadu od 8%** za poljoprivrednika van sistema PDV-a i
  prag od 8.000.000 RSD (str. 4);
- geografija referenci (Užice, Arilje, Brus, Kosjerić, Prijepolje, Priboj) je
  identična AgriX klasteru.

Drugim rečima: isti kupac, isti proizvod, ista regulativa, isti kraj Srbije.

---

## 3. Funkcionalni obim modula (iz sadržaja)

| Grupa | Stavke |
|---|---|
| Šifarnici | poljoprivredni proizvodi, otkupna mesta, ambalaža, **uslovi plaćanja (rok za uplatu)**, poslovni partneri |
| Otkup | otkup na centralnom otkupnom mestu, štampa otkupnog lista, **štampa nalepnica**, **nalog za knjiženje** |
| Ostala dokumenta | **revers za preuzetu ambalažu** na centralnom mestu |
| Izveštaji | ukupne otkupljene količine (ceo period / period / po poljoprivrednicima), preuzeto robe zbirno/pojedinačno/po datumu, vraćeno-preuzeto ambalaže |
| Kartica | kartica robe (uz **grafički prikaz** kretanja), kartica ambalaže |
| Lager lista | stanje i vrednost robe u skladištu, robe i ambalaže |
| Statistika | otkupljene količine po proizvođačima, otkupljeno u periodu, ambalaža preuzeto/vraćeno, **pivot tabele ambalaže** |
| Finansije | **kartica poljoprivrednog proizvođača** (dugovanje) → dugme **Virman** → nalog za prenos + štampa |
| Poslovne knjige | **KEP knjiga**, automatsko proknjižavanje dokumenata |
| Prečice | toolbar za tri šifarnika |

---

## 4. Matrica poređenja

Legenda: ✅ ima · ⚠️ delimično / drugačije · ❌ nema · „—" nije opisano u uputstvu

### 4.1. Otkup i obračun

| Funkcija (SOFTEK, strana) | SOFTEK | AgriX | Dokaz na AgriX strani |
|---|:--:|:--:|---|
| Otkupni list u dva TAB-a: zaglavlje (datum, otkupno mesto, poljoprivrednik, uslovi plaćanja) → stavke (str. 10–11) | ✅ | ✅ | `frmOtkup` + `modOtkup`, blokovi (`modOtkupBlok`, `clsBlokUI`) |
| **Neto = bruto − (kom ambalaže × jed. težina)**, program računa automatski (str. 11) | ✅ | ✅✅ | `frmOtkup` linije 938–974: isti obračun, uz zaštitu „tara ≥ bruto" (`DOK_MSG_TEZINA_AMBALAZE`) i **odvojeno za Klasu I i Klasu II**; težina iz `tblTipAmbalaze.TezinaGajbiceKg` |
| Fakturna cena = nabavna cena sa PDV-om iz šifarnika (str. 11) | ✅ | ⚠️ | AgriX ima dva modela cene: `tblArtikli.CenaPoJedinici` i append-only `tblCenovnik` (`GetVazecaCena`) — bogatije, sa istorijom |
| Ambalaža na otkupnom listu: `ZADUŽENO / RAZDUŽENO / STANJE` (str. 12) | ✅ | ✅ | `modAmbalaza` ledger, `ReportKarticaAmbalaze` |
| **Šifarnik uslova plaćanja** kao kodirana lista koja se bira na otkupnom listu (str. 8) | ✅ | ⚠️ | AgriX ima `OTKUP_ROK_ISPLATE` i `OTKUP_KLAUZULA` kao **jednu globalnu vrednost** (`modPrint`), ne šifarnik po dogovoru |
| Odvojen prostor šifara: šifre proizvoda i ambalaže se ne smeju poklapati (str. 5) | ✅ | ✅ | odvojene tabele `tblArtikli`, `tblKulture`, `tblTipAmbalaze` |
| Polje „Sistem PDV-a" na partneru (šifra 3 = registrovani poljoprivrednik), JMBG + matični broj gazdinstva odvojeno (str. 9) | ✅ | ⚠️ | AgriX ima BPG i PIB/JMBG polja; status u sistemu PDV-a kao kodirano polje treba proveriti |
| PDV nadoknada 8%, ukalkulisana u cenu (str. 4, 15) | ✅ | ✅ | `CFG_PDV_NADOKNADA_STOPA`, default 8 (`modConfig`) |
| Dve klase na jednom otkupu | — | ✅ | `chkDveKlase` u `frmOtkup` |
| Otkup sa terena preko mobilnog uređaja | — | ✅ | PWA `src/js/features/otkup/`, offline sync |

### 4.2. Dokumenti i štampa

| Funkcija (SOFTEK, strana) | SOFTEK | AgriX | Dokaz |
|---|:--:|:--:|---|
| Štampa otkupnog lista (str. 12) | ✅ | ✅ | `modPrint`, `OTKUP_PRINT_MODE`, PWA `otkup/otkupni-list.js` |
| 🟢 **Štampa nalepnica** — broj nalepnica = `neto / 12 + 1`; na nalepnici: firma, proizvod, datum otkupa, naziv iz šifarnika, poljoprivrednik + trocifrena šifra (str. 13–14) | ✅ | ❌ | AgriX nema štampu nalepnica; „etiketa" u `modPaletniList` je logički identitet palete (RELABEL), ne odštampana nalepnica |
| 🟡 **„Štampaj kao priznanicu"** — isti dokument, drugi ispis (str. 13) | ✅ | ⚠️ | AgriX ima print-mode po dokumentu, ne alternativni pravni oblik istog dokumenta |
| 🟢 **Nalog za knjiženje** sa kontima — 1311 magacin robe, 287 PDV plaćen poljoprivredniku, 435 dobavljači u zemlji (str. 15) | ✅ | ❌ | AgriX nema kontni nalog ni izlaz ka knjigovodstvu |
| **Revers za preuzetu ambalažu**, pravi se **pre početka otkupa** za svakog proizvođača (str. 16–18) | ✅ | ✅ | `modAmbalaza`; AgriX vodi i palete (`tblPaleta`, `modPaletniList`) |
| Otpremnice, zbirne, fakture, SEF | — | ✅ | `modDokumenta`, `modFaktura`, `modSEF*` |

### 4.3. Pregledi, kartice i lager

| Funkcija (SOFTEK, strana) | SOFTEK | AgriX | Dokaz |
|---|:--:|:--:|---|
| Ukupne otkupljene količine: ceo period / period / po poljoprivrednicima (str. 19–21) | ✅ | ✅ | `ReportOtkupRoba`, `ReportOtkupListe`, `ReportProsecnaCena` |
| Preuzeto robe zbirno / pojedinačno / po datumu; vraćeno-preuzeto ambalaže (str. 22) | ✅ | ✅ | `ReportAmbalaza`, `ReportKarticaAmbalaze` |
| Kartica robe — sve promene po artiklu (str. 23) | ✅ | ✅ | `ReportKarticaRobaRekap` |
| 🟡 **Grafički prikaz kretanja proizvoda** uz karticu (str. 24) | ✅ | ❌ | AgriX nema grafikone ni u VBA ni u PWA izveštajima |
| Kartica ambalaže (str. 25) | ✅ | ✅ | `ReportKarticaAmbalaze`, `PrintKarticaAmbalazePDF` |
| Lager lista robe/ambalaže — stanje **i vrednost** u skladištu (str. 26) | ✅ | ⚠️ | `TBL_LAGER` je deklarisan a nekorišćen; stanje se vodi kroz palete/prijemnice, ne kroz jednu lager listu sa vrednošću |
| 🟡 **Pivot tabele** preuzete/vraćene ambalaže (str. 30) | ✅ | ⚠️ | AgriX živi u Excelu, pa je pivot dostupan ručno — ali nije ponuđen kao gotov izveštaj |
| Marža, saldo kupaca, sledljivost | — | ✅ | `modMarza`, `ReportSaldoKupci`, `modSledljivost` |

### 4.4. Finansije i knjige

| Funkcija (SOFTEK, strana) | SOFTEK | AgriX | Dokaz |
|---|:--:|:--:|---|
| Kartica poljoprivrednog proizvođača — obaveze prema dobavljaču (str. 31) | ✅ | ✅ | `ReportKarticaKooperanta`, `PrintKarticaPDF` |
| 🟢 **Virman iz kartice** — jedan klik → nalog za prenos → štampa (str. 31) | ✅ | ⚠️ | AgriX ima jače: `GenerisiNalogeCSV` (CSV za e-banking, poziv na broj = broj bloka) + `PrintIsplataSpecifikacija`; ali **nema štampu pojedinačnog virmana iz kartice** |
| 🟢 **KEP knjiga** — automatsko vođenje, štampa (str. 33) | ✅ | ❌ | AgriX nema KEP ni poslovne knjige |
| Automatsko proknjižavanje dokumenata u poslovne knjige (str. 33) | ✅ | ❌ | AgriX se zaustavlja na fakturi i SEF-u |
| Uvoz bankarskih izvoda i auto-mapiranje uplata | — | ✅ | `modBankaImport` + 4 parsera, `modBankaMapiranje` |
| Avansi, kes isplate, storno lanac | — | ✅ | `ApplyAvansToOtkup_TX`, `CFG_KES_ISPLATE`, `modStorno*` |

---

## 5. Sažetak: šta SOFTEK ima a AgriX nema

1. 🟢 **KEP knjiga i nalog za knjiženje sa kontima** (str. 15, 33). Ovo je najveća
   razlika i ujedno objašnjenje SOFTEK-ovog ugla: otkup je modul na knjigovodstvenoj
   platformi. Kupac koji hoće „i otkup i knjige na jednom mestu" tamo dobija oboje.
2. 🟢 **Štampa nalepnica** sa formulom `neto / 12 + 1` (str. 13–14). Sitno, vidljivo,
   svakodnevno.
3. 🟢 **Virman/nalog za prenos direktno iz kartice proizvođača** (str. 31).
4. 🟡 **Grafički prikaz** kretanja artikla (str. 24) i **pivot tabele** ambalaže (str. 30).
5. 🟡 **Šifarnik uslova plaćanja** umesto jedne globalne klauzule (str. 8).

`INFERENCE`: nijedna od ovih stavki nije arhitektonska. Sve su u dometu AgriX-a; KEP i
kontni nalog su jedini koji traže poslovnu odluku (da li AgriX ulazi u knjigovodstvo
ili se namerno zaustavlja na SEF-u).

## 6. Sažetak: šta AgriX ima a u uputstvu ga nema

Cela desna strana lanca i ceo teren:

| Oblast | AgriX |
|---|---|
| Prodaja | kupci, otpremnice, zbirne, fakture, **SEF** sa state machine-om |
| Teren | PWA za kooperanta, vozača, otkup i management; offline, QR, potpisi |
| Logistika | vozači, dispečer, transport, palete, paletni list, hladnjača, prerada |
| Finansije | uvoz izvoda 4 banke, auto-mapiranje, CSV nalozi, avansi, marža |
| Agronomija | parcele sa GIS poligonima, GGAP, knjiga polja, agromere, karenca, agrohemijski magacin i dug |
| Kontrola | storno centar sa žurnalom i recovery, sledljivost lanca, monitoring, prava po 12 oblasti |
| Isporuka | self-update, licenciranje po uređaju, health check |

`LIMITATION`: uputstvo pokriva samo modul otkupa — SOFTEK kao firma verovatno ima i
fakturisanje i SEF u drugim modulima. Ne tvrditi suprotno bez dokaza.

## 7. Prodajna implikacija

`INFERENCE`:

1. **SOFTEK je referentna tačka za cenu i za „šta je dovoljno".** Kupac koji ga koristi
   ima pokriven otkup, ambalažu, karticu i knjige — dakle nije u bolu zbog osnovnog
   dokumenta. AgriX se ne prodaje protiv otkupnog lista, nego protiv svega što SOFTEK
   modul ne dodiruje: terena, hladnjače, transporta, banke i SEF-a.
2. **Dve male stvari mogu da odluče demo** — nalepnice i „gde su mi knjige". Vredi imati
   spreman odgovor na oba (nalepnice: da/ne/kada; knjige: AgriX izvozi ka knjigovodstvu,
   ne vodi ih).
3. **Preklapanje geografije je potpuno** (Zapadna i Centralna Srbija), pa je verovatnoća
   susreta u istom poslu visoka — viša nego sa AGROSOFT-om.

## 8. Šta ostaje da se proveri

1. Puni obim SOFTEK proizvoda — da li postoje moduli za prodaju, fakturisanje i SEF.
2. Cena i model održavanja.
3. Da li postoji mobilna ili web komponenta (uputstvo iz 2017. je ne pominje).
4. Status 13 referenci iz `competitor_references.csv` — koliko ih danas radi.
5. Da li se SOFTEK pojavio u nekom AgriX poslu → ako jeste, red u `competitive_events.csv`.

`DECISION` (predlog, bez ID-ja): SOFTEK se vodi kao **direktan konkurent** — za razliku
od AGROSOFT-a koji je adjacent. Ako se usvoji, treba mu dodeliti ID u
`docs/Master Plan/09_QA_DECISION_LOG.md` i upisati ga u `05_COMPETITION.md`.
