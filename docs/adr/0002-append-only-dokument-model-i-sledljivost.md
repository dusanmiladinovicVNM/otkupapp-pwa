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

### 3. Srce — 3 in-place rekalka → storno+novi (`modDokumentInvariant`)
`RecalculateZbirnaFromOtpremnice_TX`, `RecalcOrStornoEmptyZbirna_TX`, `ApplyKlasaRecalc`:
umesto upisa na postojeći zbirna red → **storno tekućih zbirna redova (BrojZbirne) +
append novih** sa preračunom + `IspravkaOd`/`CorrectionID`. Deca (isti BrojZbirne)
se ne diraju.

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
- **PWA sync (smer a)** NIJE prost `ZbirnaID -> BrojZbirne` swap: `RequireSingleMasterSyncRow`
  trazi tacno jedan red, a `BrojZbirne` daje 2 (klase). Migrira se na
  `(broj, klasa)` -> aktivan red.
- **Svi citaci** biraju tekucu verziju = aktivan red po `(broj, klasa)`.
- **Primitiv:** `modDokumentInvariant.FindSingleActiveRow(tbl, brojCol, broj, klasaCol,
  klasa)` -> indeks jedinog aktivnog reda; 0 = nema; -1 = vise aktivnih (integritet
  povreda; u append-only sme najvise jedan aktivan po `(broj, klasa)`). (Faza 7 korak 3.0.)

## Predložene faze (Faza 7 / v3) — stanje

1. ✅ Šema (trace kolone) — non-breaking. *(korak 1)*
2. ✅ Utiskivanje `IspravkaOd`/`ZamenjenSa`/`CorrectionID` na ISPRAVKA. *(korak 2)*
4. 🟡 Sledljivost vidljiva: panel `[ispravka dokumenta X]` + štampa otpremnice. *(korak 4, deo)*
3. Korak 3 (append-only zbirne), smer (a), pod-koraci:
   - **3.0** `FindSingleActiveRow` `(broj, klasa)` → aktivan red + test. *(u toku)*
   - **3.1** PWA sync (`modMasterSync`) `ZbirnaID` → `(broj, klasa)`-aktivno. *(eksterno; test pre)*
   - **3.2** `RecalculateZbirnaFromOtpremnice_TX` → storno tekućih `(broj,*)` + append novih.
   - **3.3** Čitači → aktivna verzija po `(broj, klasa)`.
5. `blok → BrojOtpremnice`; Izdato-status kapija; penzionisati in-place + gard-e.
6. Testovi (`modTestStornoCentar`) + ADR-0002 → „Prihvaćeno" + ARCHITECTURE_REFERENCE.

## Alternativa (odbijena)

„Ostaviti in-place uz gard-e" — odbijeno: gard-i leče simptom, ne daju sledljivost
unazad; i dalje menjaju izdate dokumente. Puni event-sourcing rewrite iz nule —
odbijeno kao nepotreban (model je već ~70% postavljen).
