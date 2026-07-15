# Storno centar — backlog (zapamćeni predlozi za dalji rad)

> Konsolidovani predlozi iz rada na grani `claude/storno-dedup-ux-k6esuw`.
> Održavati kao živu listu. Prioritet: P1 (pre šireg merge-a) → P3 (kasnije).

## P1 — pre šireg merge-a / korektnost

- **[TESTS] `modTestStornoCentar` — automatske rollback-safe asertacije.**
  Bar: undo-garda (`OtkupBlockDeadParent`/`UndoStorno_TX`), Guard C
  (`BlockStornoDriftReason`), impact agregator (`BuildStornoImpact`,
  `GetPaleteImpactByField`). Obrazac kao `modTestStorno`/`modTestPalete`
  (fixture → poziv → `Debug.Assert` → rollback). Bez UI-a.
  *(zamerka #1 iz code-review-a; W1.3 nikad slelo automatski)*

## P2 — arhitektura / profesionalizacija (ADR-0002)

- **[ARCH] Append-only dokument model + sledljivost.** Vidi `docs/adr/0002-...`.
  Potpuno eliminisati in-place; storno+novi red svuda; `IspravkaOd`/`CorrectionID`
  utisnut na dokument. Fazno (Faza 7): šema → `blok→BrojOtpremnice` → agregati
  storno+novi → čitači/print/PWA → izdato-status kapija → testovi.
- **[ADR-0001] Izdato/prosleđeno stanje** (`IzdatoStatus`) — preduslov da se granica
  interni/prosleđen predstavi u podacima (tačka B).
- **[ADR-0001] In-place rekalk zbirne samo dok nije prosleđena** (tačka A); posle →
  storno+reizdaj + korektivni dokument ka kupcu. Najoštrije u
  `CompleteOtpremnicaIspravka` (ista zbirna) i `CompleteZbirnaIspravka` (isti broj).
- **[ADR-0001] Blok-atributna izmena** (ime/klasa/sorta) bez diranja količine —
  zaseban put, NE storno bloka.

## P2 — Faza 6 (iz plana rada)

- **[W6.1] Otkup/Faktura/Novac u framework** (trenutno prost storno).
- **[W6.2] Otkupni list pri stornu bloka** — auto-reprint „STORNIRANO" (sad samo
  označen).
- **[W4.3] Vidljiv baner „ISPRAVKA u toku"** dok operater unosi novi dokument
  (sad samo `m_activeCorrectionID` module-level trag).
- **[W3.1] Prelij/preko palete inline** u panelu (sad modalni `PaletaAdjustPrompt`).
- **[W3.3] PONIŠTENJE potvrda inline** (sad modalni MsgBox `res("blocked")`).

## P3 — održivost / polish

- **[REFACTOR] `frmDokumenta.frm` bloat** (6062 linije). Izneti Storno-centar overlay
  u `clsStornoCentarUI` (WithEvents klasa, bez `.frx` shell-a) ili namensku
  `frmStornoCentar` (traži jednokratni dizajnerski korak). *(zamerka #2)*
- **[CLEANUP] Mrtav kod:** `ShowNedovrsenoPanel` (count-only, penzionisan),
  `SetupStornoPregledButton`/`SetupRecoveryButton`/`SetupNedovrsenoButton`,
  `m_btnNedovrseno`/`m_btnStornoFind` — više se ne pozivaju posle konsolidacije.
- **[UI] Bold naziv moda + opis** iznad dugmadi; opciono blaga pozadina u boji dugmeta.
- **[PWA] KI-006** `ExportMagacinKoop` ne izuzima `ART_POCETNI_DUG` (poznato, van storna).

## Urađeno na ovoj grani (referenca)

- Faze 0–5 (uz odložene W3.1/W3.3/W4.3), konsolidacija ulaza (sve kroz Storno),
  upozorenja u pregledu, ujednačen „Efekat storna" (Duplikat pa Poništenje),
  idle pre-warm keš (`modStornoWarm`), naziv partnera umesto ID.
- Guard #3 (undo-siroče, `OtkupBlockDeadParent`).
- Guard C (blok-drift nad živom otpremnicom, `BlockStornoDriftReason`).
- ADR-0001 (nepromenljivost izdatih dokumenata), ADR-0002 (append-only model, predlog).
