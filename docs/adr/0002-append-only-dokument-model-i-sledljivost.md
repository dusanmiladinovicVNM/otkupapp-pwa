# ADR-0002 — Append-only dokument model + sledljivost ispravki (bez in-place izmena)

- **Status:** Predloženo (analiza smera; NIJE implementirano)
- **Datum:** 2026-07-13
- **Gradi na:** ADR-0001 (nepromenljivost izdatih dokumenata)
- **Cilj:** potpuno eliminisati in-place izmene u lancu dokumenata. Svaka korekcija =
  **storno starog reda + kreiranje novog reda**, gde novi red **vidljivo nosi da je
  ispravka starog** (prava sledljivost unazad: „zašto je nešto urađeno").

## Odluka (smer)

1. **Dokument-red je nepromenljiv posle izdavanja.** Ne edituje se; menja se storno
   (`Stornirano=Da`) + **novi red**.
2. **Novi red nosi referencu na stari:** `IspravkaOd` (stari broj) + `CorrectionID`
   (veza na `tblStornoVeze`). Stari red dobija `ZamenjenSa` (unapred pokazivač).
3. **Poslovni ključ je identitet, red je verzija.** Deca referišu **poslovni ključ**
   (BrojZbirne, BrojOtpremnice, BrojPrijemnice), nikad row-ID → re-verzija roditelja
   NE zahteva prevezivanje dece.
4. **Agregati se ne rekalkulišu u mestu** — storno starih redova + novi redovi.
5. **Palete/operativno stanje su IZUZETAK:** paletni header (kg/amb/gajbe) je
   **izvedena projekcija** iz aktivnih stavki, ne izdati dokument → računa se
   (ne verzioniše kao dokument). Paletne **stavke** ostaju append-only (storno+nova).

## Zašto je kod već ~70% spreman (bitno za procenu)

- **Soft-delete** (`Stornirano`) svuda — osnova append-only-a.
- **Deca referišu poslovni ključ** (BrojZbirne/BrojPrijemnice) — JEDINI izuzetak je
  `blok → otpremnica` preko **row-ID** (`COL_OTK_OTPREMNICA_ID`); blok ionako nosi i
  `BrojZbirne`.
- **`tblStornoVeze` VEĆ ima punu sledljivost** (Old/New/Parent DocType+ID+Broj, Mode,
  Status, CreatedAt/By, Message…). „Novi→stari" veza već postoji — samo nije
  „utisnuta" na sam dokument-red i ne pravi se za svaku verziju.
- **Append-only presedani postoje** (ambalaza ledger, cenovnik).
- **`LookupActiveID` / `ExcludeStornirano`** = već „uzmi tekuću aktivnu verziju".

Zato ovo NIJE event-sourcing rewrite iz nule, nego **usmereno dovršavanje** modela
koji već postoji.

## Šta i koliko mora da se menja (change surface)

### 1. Šema (malo, non-breaking) — `modSetup Ensure*Schema`
Dodati na dokument-tabele (`tblOtkup, tblOtpremnica, tblZbirna, tblPrijemnica,
tblFaktura, tblNovac`): `IspravkaOd`, `ZamenjenSa`, `CorrectionID`, `IzdatoStatus`
(+ opciono `Verzija`). Idempotentna migracija + backfill (istorija = `IzdatoStatus=IZDATO`).

### 2. Jedna referenca (srednje) — `blok → otpremnica` po **BrojOtpremnice**
Umesto `OtpremnicaID` (row-ID). Time re-verzija otpremnice ne dira blokove. Dodiruje:
otkup save, `GetBlokOtkupIDs`, `ActiveBlocksForFlow`, `GetOtpremnicaIDsByBroj`,
`FreeOtkupBloksInline`, `ReassignOtkupToOtpremnica_TX`.

### 3. Srce — auto-recalk zbirne (`modDokumentInvariant`) → AUDIT, ne re-verzija (revidirano 2026-07-15)

> **Revizija posle A0 analize sync-a (2026-07-15).** Prvobitna zamisao (storno tekućih
> zbirna redova + append novih **sa istim BrojZbirne**) je **odbačena** kao nebezbedna,
> i sama ideja re-verzionisanja svakog auto-recalk-a je preispitana. Original je bio:
> *„umesto upisa na postojeći zbirna red → storno + append novih sa preračunom"*.

**Nalazi A0 (dokazano, file:line u analizi):**
1. **Sync je bezbedan na nov `ZbirnaID`.** `ZbirnaID` je čisto master-interni — nikad
   ne prelazi sync granicu; ceo lanac (PWA⇄GAS⇄master) ključa na `ClientRecordID`.
   Nov `ZbirnaID` za istu zbirnu je sync-u nevidljiv (nema duplikata/gubitka veze).
2. **ALI isti `BrojZbirne` na dva aktivna reda je nebezbedan.** `BrojZbirne` je interni
   join-ključ celog lanca (otpremnica→zbirna, otkup-blok lookup, recovery, izveštaji);
   dva živa reda istog broja se sudaraju. `StampIspravkaTrace` k tome **no-op** kad
   `newBroj == oldBroj`.
3. **Eksplicitna ISPRAVKA zbirne VEĆ radi pravilan append-only** (`CompleteZbirnaIspravka`):
   nov `BrojZbirne` + relink otpremnica/prijemnica + rekalk + validacija + trace.
4. **Auto-recalk je izveden agregat** (zbirna = suma aktivnih otpremnica), ne autorska
   korekcija. Puno re-verzionisanje pri svakom recalk-u = ogroman churn (broj zbirne se
   menja pod nogama operatera, sva deca se relinkuju) bez poslovne vrednosti.

**Odluka:** `RecalculateZbirnaFromOtpremnice_TX` / `ApplyKlasaRecalc` **ostaju in-place**
(izveden agregat), ali kad recalk promeni **IZDATU** zbirnu (`DocIsIssued`) → upisuje
se **audit-trag** (`Monitor_Event ZBIRNA_IZDATA_RECALC`, WARN): `CorrectionID` + razlog +
`stara→nova` (kg/amb) po klasi. Tako izmena izdatog dokumenta ostavlja trag (ADR-0001)
bez nebezbednog re-verzionisanja. Eksplicitne korekcije i dalje idu kroz
`CompleteZbirnaIspravka` (pravi append-only, nov broj). Test: `Test_ZbirnaRecalcInPlace_Auto`.

### 4. Relinkovi — VEĆINA NESTAJE
Sa stabilnim ključem + append-only, re-verzija ne traži relink. Ostaje samo pravi
**re-parenting** (malina 1:1: prijemnica na NOVU zbirnu) → to se svede na storno+reizdaj
pomerenog dokumenta, ne tihi `BrojZbirne` edit. `ReassignOtkupToOtpremnica_TX` →
eliminisan (blok čuva broj). `FreeOtkupBloks*` (unbind) → odluka: status-prelaz ili
storno+nova „čeka" verzija.

### 5. Palete — izvedeno, ne dokument
`DecrementPaletaForStavka`, `AdjustPaletaGajbiceZaPrijemnicu_TX`, `PaletaAdjustPrompt`
→ „preračunaj header iz AKTIVNIH stavki" (derivacija). Stavke: append-only (storno+nova).

### 6. Čitači — uvek tekuća aktivna verzija po ključu (široko, mehanički)
Svi izveštaji/agregati/print/PWA/faktura/novac: `LookupActiveID`/`ExcludeStornirano`.
Više storniranih verzija + jedna aktivna po broju → svuda birati aktivnu.

### 7. Sledljivost utisnuta na red + na print/PWA
`IspravkaOd`/`CorrectionID` na dokument-redu. Print: baner „Ispravka dokumenta X od Y".
Sync (`modStammdatenSync`/`modMasterSync`): prenosi `IspravkaOd`/`Stornirano`/`IzdatoStatus`.

### 8. Izdato-status kapija (ADR-0001 tačka B)
`IzdatoStatus` (DRAFT/IZDATO/PROSLEDJENO): draft → slobodna izmena; izdato → obavezno
storno+novi; prosleđeno → i korektivni dokument ka kupcu.

**Nalaz (2026-07-13):** ova app **NEMA draft-fazu** za chain dokumente — svaki je IZDAT
čim se snimi (grep `draft/nacrt` = samo SEF status fakture, nevezano). Zato je
`IzdatoStatus` podrazumevano IZDATO; **prazno = IZDATO** (konzervativna konvencija, bez
backfill-a). Gate primitiv: `modDokumentInvariant.DocIsIssued(tbl, brojCol, broj)` (True
osim ako je eksplicitno DRAFT) + `SetIzdatoStatus`. DRAFT rezervisan (buduci parkiran
dokument), PROSLEDJENO za buduci sync-push. Posledica: in-place se NE gejtuje uslovno —
uvek append-only (Korak 3.2), bez „draft izuzetka".

### 9. Penzionisati legacy in-place putanje + gard-e
Kad je storno+novi svuda, drift ne postoji → Guard C / undo-garda (#3) se pojednostave.

### 10. Testovi + docs
`modTestStornoCentar`: invarijante (nijedan izdati red se ne menja; sledljivi lanac
`IspravkaOd` ceo; tekuća-verzija selekcija; `zbirna = Σ aktivnih otpremnica`).
Ažurirati `ARCHITECTURE_REFERENCE`, runbook.

## Je li ovo korak u pravom smeru? — DA

- Standardni računovodstveni/ERP model: nepromenljiva knjiženja + storno/korektivna
  knjiženja sa referencom.
- **Prava sledljivost unazad** (`IspravkaOd`/`CorrectionID` lanac → „zašto").
- **Rešava ADR-0001 u korenu** (A: in-place rekalk, B: izdato-flag, C: blok-drift) —
  gard-i postaju nepotrebni.
- **Pojednostavljuje motor** (nestaje „recalc-u-mestu vs zameni" dilema i većina
  relinkova) jer se oslanja na već-postojeće poslovne ključeve.

## Rizici i mitigacije

- **Široki čitač-audit (tačka 6)** — mehanički ali obiman; mitigacija: standardizovati
  na `LookupActiveID`, pa test-invarijante.
- **Rast broja redova** (verzije) — beznačajno za veličine ove app; `ExcludeStornirano`
  kes već postoji.
- **Big-bang rizik** — NE raditi odjednom; fazno (dole).
- **Over-engineering paleta** — svesno izuzete (tačka 5).

## Kanonsko adresiranje: `(broj, klasa)` -> AKTIVAN red (dopuna, 2026-07-13)

Svaki dokument lanca (Otkup/Otpremnica/Prijemnica/Zbirna) je **linijski model**:
jedan poslovni broj ima **VISE redova, jedan po klasi** (Klasa I / II). To NIJE
greska modela nego standardni „document lines as rows" obrazac; za append-only je
i **prednost** (storno jedne klase = storniraj taj red + append novi; „siroki"
model sa fiksnim KolicinaI/II kolonama bi trazio parcijalni in-place edit deljenog
reda -> bori immutability).

Posledica: identitet reda nije `broj` (vise redova) ni `RowID` (menja se po verziji),
nego **kompozitni poslovni kljuc `(broj, klasa)` razresen na AKTIVAN (ne-stornirani)
red**. Jedinstven je **medju aktivnim redovima** (v1-stornirano + v2-aktivno dele
`(broj, klasa)`).

Zato:
- **Svi citaci** biraju tekucu verziju = aktivan red po `(broj, klasa)`.

### PWA/sync je VEC po BrojZbirne (uvid u `src/js/` + `modMasterSync`, 2026-07-13)

Provera PWA koda (isti repo) je pokazala da je **cross-system ugovor `BrojZbirne`
(poslovni kljuc), ne `ZbirnaID`**:
- PWA **generise `brojZbirne` client-side** (`src/js/features/vozac/zbirna.js`), salje
  ga, backend ga vraca; PWA radi sa `klasa` (I/II/"I+II"). Otpremnica-transport isto
  po `brojZbirne`.
- Excel link otpremnica->zbirna ide preko **`BrojZbirne`** (`LinkOtpremnicaToBrojZbirneStrict`);
  komentar u `modMasterSync` izricito: *"ValidateZbirna/prijemnica/faktura vezu drze
  preko BrojZbirne."*
- `RequireSingleMasterSyncRow` se koristi **samo sa row-ID kolonama** (OtkupID/OtpremnicaID/
  ZbirnaID = jedinstveni), NIKAD sa `BrojZbirne` -> strah "2 klase -> RequireSingle pada"
  NE vazi za sync.
- Jedini `ZbirnaID`-ulaz (`GetBrojZbirneForIDStrict`) je **tranzijentan** (Excel tek
  mintovao ID pa mu trazi broj) i resava se preko reda koji ostaje (storniran) -> broj
  je stabilan.

**Posledica:** smer (a) **ne trazi PWA izmenu** (ugovor je vec `BrojZbirne`), a
re-verzija zbirne (novi `ZbirnaID`, isti `BrojZbirne`) **ne lomi sync** jer sve veze
idu preko `BrojZbirne`. Korak 3.1 se svodi na Excel-stranu proveru da citaci uzimaju
AKTIVNU verziju (`FindSingleActiveRow`), a ne na rizicnu eksternu migraciju.
- **Primitiv:** `modDokumentInvariant.FindSingleActiveRow(tbl, brojCol, broj, klasaCol,
  klasa)` -> indeks jedinog aktivnog reda; 0 = nema; -1 = vise aktivnih (integritet
  povreda; u append-only sme najvise jedan aktivan po `(broj, klasa)`). (Faza 7 korak 3.0.)

## Predložene faze (Faza 7 / v3) — stanje

1. ✅ Šema (trace kolone) — non-breaking. *(korak 1)*
2. ✅ Utiskivanje `IspravkaOd`/`ZamenjenSa`/`CorrectionID` na ISPRAVKA. *(korak 2)*
4. 🟡 Sledljivost vidljiva: panel `[ispravka dokumenta X]` + štampa otpremnice. *(korak 4, deo)*
3. Korak 3 (append-only zbirne), smer (a), pod-koraci:
   - **3.0** ✅ `FindSingleActiveRow` `(broj, klasa)` → aktivan red + rollback-safe test.
   - **3.1** ~~PWA migracija~~ → **nije potrebna**: sync je već po `BrojZbirne` (vidi gore). Svodi se na Excel-stranu proveru čitača (deo 3.3).
   - **3.2** `RecalculateZbirnaFromOtpremnice_TX` → storno tekućih `(broj,*)` + append novih (+ `IspravkaOd`). *Core; test pre.*
   - **3.3** Čitači → aktivna verzija po `(broj, klasa)` (`FindSingleActiveRow`).
5. `blok → BrojOtpremnice`; Izdato-status kapija; penzionisati in-place + gard-e.
6. Testovi (`modTestStornoCentar`) + ADR-0002 → „Prihvaćeno" + ARCHITECTURE_REFERENCE.

## Alternativa (odbijena)

„Ostaviti in-place uz gard-e" — odbijeno: gard-i leče simptom, ne daju sledljivost
unazad; i dalje menjaju izdate dokumente. Puni event-sourcing rewrite iz nule —
odbijeno kao nepotreban (model je već ~70% postavljen).
