# ADR-0001 — Nepromenljivost izdatih dokumenata; korekcija = storno + reizdaj

- **Status:** Prihvaćeno (princip); sprovođenje delimično — vidi „Usklađenost".
- **Datum:** 2026-07-13
- **Kontekst grane:** `claude/storno-dedup-ux-k6esuw` (Storno centar)
- **Vezano:** `docs/production-runbook-document-chain.md`, `docs/INTEGRITET_PROVERE.md`,
  `docs/STORNO_CENTAR_PLAN_RADA.md`, `modStornoFlow`, `modDokumentInvariant`, `modStorno`.

## Kontekst

Lanac dokumenata: `Otkup (blok) → Otpremnica → Zbirna → Prijemnica → Faktura → Novac`.
Invarijanta količine: `otpremnica = Σ svojih blokova`, `zbirna = Σ svojih otpremnica`.

Kad je blok pogrešan, prirodan poriv je „storniraj taj blok i unesi ispravan; do
tada je disbalans, unosom se vraća ravnoteža". To je ispravan **workflow princip**
(ispravka = privremena neravnoteža → zamena → ravnoteža), ali oslanja se na
pretpostavku da će sistem sam da re-balansira.

Presudno ograničenje: kada je dokument **izdat i prosleđen** (otpremnica/zbirna
kod kupca, vezana za njegovu prijemnicu), on je **zamrznut**. Druga strana drži
kopiju sa tim brojem. Tiha izmena naše verzije (direktno ili preko storna jednog
bloka koji obara zbir) razilazi naše i kupčeve podatke → neuparivo, spor, pogrešna
faktura. Posle slanja to više nije interna nekonzistentnost, nego **divergencija
od onoga što kupac drži**.

## Odluka

1. **Izdat + prosleđen dokument je nepromenljiv.** Ne edituje se — ni direktno, ni
   posredno preko storna pojedinačnog bloka koji bi mu promenio zbir.

2. **Korekcija se radi NOVIM dokumentom, ne izmenom.**
   - Stari dokument se **STORNIRA** (ostaje zamrznut u audit tragu, `Stornirano=Da`).
   - Izda se **ISPRAVAN novi** dokument.
   - Kupac dobija **korektivni dokument koji referiše original** (broj + razlog),
     da preveže svoju prijemnicu. Veza se čuva **referencom**, nikad tihom
     promenom broja.

3. **„Header = Σ stavki" važi nad AKTIVNIM skupom.** Rekalk ne znači „izmeni
   zamrznut dokument", nego „preračunaj živi agregat nad novim aktivnim skupom
   pošto si stari zamenio korektivnim dokumentom".

4. **Blok nije samostalno ispravljiv unutar izdatog lanca.** Storno pojedinačnog
   bloka koji ima žive uzvodne dokumente NE sme tiho da ostavi otpremnicu/zbirnu u
   neskladu. Takav zahtev se **preusmerava na dokument-level ISPRAVKA** (storno cele
   otpremnice + reizdaj), ne rešava se in-place rekalkom izdatog dokumenta.

5. **Granica odluke = stanje dokumenta (interni/draft vs izdat/prosleđen).**
   - Dok dokument **nije izdat/prosleđen** → rekalk u mestu je dozvoljen (nema
     spoljnog uticaja).
   - Čim je **izdat/prosleđen kupcu** → isključivo storno + reizdaj + korektivni
     dokument ka kupcu.

6. **Nikad tihi drift.** Ako je privremeni disbalans neizbežan (ispravka u toku),
   on mora biti **prvoklasna, praćena stavka** (Nedovršeno + health-check), sa
   vlasnikom i putem do zatvaranja — nikad tiho u živom agregatu.

## Posledice

- **Modovi zadržavaju čisto značenje:**
  - *ISPRAVKA* = storno cele + reizdaj (privremeni, praćen disbalans do zamene).
  - *DUPLIKAT / PONIŠTENJE* = količina je stvarno nestala → automatski rekalk /
    kaskadni storno; bez „ostavi uz upozorenje".
  - *Atributna greška (ime/klasa/sorta), količina ista* = izmena bez diranja
    količine (agregati nepromenjeni) — NE storno bloka.
- **Potreban je signal „izdato/prosleđeno".** Bez njega sistem ne može da razlikuje
  internu izmenu od izmene tuđeg dokumenta (vidi Usklađenost, tačka B).
- **Blok-level storno nad izdatim lancem se zabranjuje/preusmerava**, ne dobija
  in-place rekalk.

## Usklađenost trenutnog koda (2026-07-13, ova grana)

Snimak stanja u trenutku donošenja odluke:

- **A. In-place rekalk zbirne postoji.** `modDokumentInvariant.RecalculateZbirnaFromOtpremnice_TX`
  i `RecalcOrStornoEmptyZbirna_TX` **upisuju novu količinu na POSTOJEĆI zbirna red**
  (isti broj), a ne storno+reizdaj. Bezbedno **samo** dok je zbirna interna.
  Odstupa od Odluke #2/#5 ako je zbirna prosleđena.
- **B. Ne postoji „izdato/prosleđeno" stanje** na `tblOtpremnica`/`tblZbirna`
  (nema statusne kolone). Zato granica iz Odluke #5 danas **nije predstavljiva u
  podacima** — ovo je preduslov za sprovođenje.
- **C. Blok-level storno pravi tihi drift.** `modStornoFlow.StornoSelectedBlocks_TX
  → modStorno.StornoOtkup` ne rekalkuliše roditelja; kad roditelj preživi
  (prijemnica DUPLI/RESI_KASNIJE + čekiran blok) → otpremnica/zbirna precenjene
  (Odluka #4/#6 prekršena).
- **D. Dokument-level ISPRAVKA i PONIŠTENJE su načelno usklađeni** (storno cele +
  reizdaj / puna kaskada; original ostaje zamrznut) — osim in-place rekalka zbirne
  iz tačke A.

## Sledeći koraci (Faza 6 kandidati, ne blokiraju ovu granu)

1. Dodati stanje `IZDATO/PROSLEDJENO` (kolona ili `tblStornoVeze`-nivo) na
   otpremnicu/zbirnu → čini granicu #5 predstavljivom.
2. Blok-storno nad **izdatim** lancem: odbij + preusmeri na otpremnica-ISPRAVKA
   (ili blok-ISPRAVKA sa „čeka zamenu" kontekstom + rebalans po unosu).
3. In-place rekalk zbirne dozvoliti **samo** dok nije prosleđena; inače storno +
   reizdaj + korektivni dokument ka kupcu.
4. Atributna korekcija bloka (bez diranja količine) kao zaseban put.

## Alternativa (odbijena)

„Uvek rekalk u mestu" — odbijeno jer bi menjalo izdate dokumente koje kupac drži
(razilaženje kopija). „Tihi disbalans bez oznake" — odbijeno (Odluka #6).
