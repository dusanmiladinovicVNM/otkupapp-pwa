# Storno centar — detaljan plan rada

> Cilj: ceo storno (svi dokumenti + sve posledice na otkupne listove i palete)
> objediniti u **jedan tok i jedan uvid**, na jednom mestu. Od nalaženja
> dokumenta do rezultata i ostatka — bez skrivenih MsgBox-ova, bez skakanja po
> formama, bez rasutih „nedovršeno" prikaza.
>
> Grana: `claude/storno-dedup-ux-k6esuw`. Motori se **reuse-uju** (ne
> reimplementiraju); menja se orkestracija i prikaz.

---

## 0. Kontekst i zatečeno stanje

### Šta već postoji (na grani + main)
- **Motor storna** — `modStorno.bas`: `StornoOtkup/Otpremnica/Zbirna/Prijemnica/Faktura/Novac/Paleta/Prerada` (+ `*_TX` i `*ByBroj_TX`). Soft-delete, transakciono, bez MsgBox-a.
- **Framework (4 moda)** — `modStornoFlow.bas` + `modStornoContext.bas`: ISPRAVKA / DUPLI / PONIŠTENJE / REŠI KASNIJE, persistentni trag u `tblStornoVeze`. Na grani uvezana i **Prijemnica** + panel „Storno / potvrda" + multiselect blokova.
- **Palete motor** — `modPaletniList.bas`: `DetachOsirocenePaletaStavke_TX`, `ReassignPaleteToPrijemnica_TX`, `AdjustPaletaGajbiceZaPrijemnicu_TX` (prelij/preko), `DecrementPaletaForStavka`, `StornoPaleta`, `GetPaletaAggregates`, `GetPaleteInfoForPrijemnicaBroj`. UI prompt `PaletaAdjustPrompt` (`modPaletniListUI.bas`).
- **Recovery** — panel „Osiroceni dokumenti" (frmDokumenta) + „Pregled storniranih" + „Mod: Palete"; dva izvora istine (`GetPendingCorrections` vs izvedeno iz podataka).

### Šta ostaje rasuto (meta ovog plana)
Ulaz „napamet" (cmb + kucaš broj) · izbor moda kroz 3 MsgBox-a · ISPRAVKA prekida tok (panel → forma → MsgBox za palete) · posledice na palete/kapacitet i otkupne listove nevidljive pre akcije · tri odvojena „nedovršeno" prikaza · nekonzistentno po tipu (otkup/faktura/novac idu prostim Da/Ne).

---

## 1. Arhitektura ciljnog rešenja (slojevi)

```
UI (frmDokumenta, runtime overlay)   ← Storno centar: Nađi → Uvid → Razlog → Pod-odluke → Primeni → Rezultat/Ostatak
        │  poziva
Impact/agregacija (modStornoImpact)   ← NOVO: jedan model uticaja (lanac + listovi + palete+kapacitet + saldo + faktura + novac)
        │  poziva
Orkestracija (modStornoFlow / Context) ← Run*Correction, GetStorno*Rows, tblStornoVeze  (postoji, proširuje se)
        │  poziva
Motor (modStorno, modPaletniList)      ← Storno*_TX, Detach/Reassign/Adjust*  (NETAKNUTO — samo reuse)
```

**Odluka o formi:** ostajemo na **runtime overlay-ima u `frmDokumenta`** (`.frx` se ne dira; sve `Controls.Add` + `WithEvents`, kao postojeći paneli). Zaseban `frmStornoCentar` bi bio čistiji, ali traži jednokratno kreiranje shell-a u Excel dizajneru (binarni `.frx` se ne pravi iz koda) — ostaje kao opcija za kasnije.

---

## 2. Faze i work-items

Notacija po work-item-u: **Fajlovi · Nove/izmenjene procedure · Reuse · Gotovo (acceptance) · Rizik · Test**.

---

### FAZA 0 — Temelj: popravke iz code-review-a
*Cilj: čist temelj pre nadogradnje. Nizak rizik, brzo.*

**W0.1 — Poštenje context-a pri padu storna blokova**
- Fajlovi: `frmDokumenta.frm` (`ApplyCorrectionFromPanel`), po potrebi `modStornoFlow.bas`.
- Izmena: kad `StornoSelectedBlocks_TX` vrati `-1`, ne ostavljati context `COMPLETED` — pozvati `MarkCorrectionManual` sa porukom „blokovi nisu stornirani".
- Reuse: `modStornoContext.MarkCorrectionManual`.
- Gotovo: pad storna bloka → context reda je `MANUAL_REQUIRED`, vidljiv u „Nedovršeno".
- Rizik: nizak. Test: statički (grep putanje) + Excel: veštački obori (nevalidan blok) i proveri context.

**W0.2 — Perf `GetStornoBlockRows`**
- Fajl: `modStornoFlow.bas`.
- Izmena: jedan `GetTableData(TBL_OTKUP)` + `Dictionary` `OtkupID→red`, umesto 4× `LookupValue` po bloku.
- Gotovo: nema regresije u prikazu; O(N) umesto O(N·M).
- Rizik: nizak. Test: statički; Excel: panel nad zbirnom sa mnogo blokova.

**W0.3 — Guard na `Nothing`**
- Fajl: `frmDokumenta.frm` (`ApplyCorrectionFromPanel`): `If res Is Nothing Then Exit Sub` posle `DispatchCorrection`.
- Rizik: trivijalan.

---

### FAZA 1 — Impact agregator (motor uvida)
*Cilj: jedna funkcija = pun model uticaja za bilo koji dokument, uključujući palete (kg/amb/kapacitet) i otkupne listove. Testabilno bez UI-a.*

**W1.1 — Novi modul `modStornoImpact.bas`**
- Funkcija: `Public Function BuildStornoImpact(docType, broj, dokTip) As Object`
  Vraća `Scripting.Dictionary`:
  - `header`: `{tip, broj, partner, datum, kolicina}`
  - `chain`: `Collection` dict-ova `{dok, broj, actByMode{ISPRAVKA,DUPLI,PONISTENJE,RESI}={text, kind}}`
  - `blocks`: `Collection` `{otkupID, brDok, kg, klasa, koop}`
  - `palete`: `Collection` `{paletaID, label, used, cap, neto, amb, preradjena, deltaByMode{gajb,kg,amb}}`
  - `ambSaldo`: procena obrta ledgera po modu
  - `faktura`: `{hasFaktura, id, broj}`
  - `novac`: `{iznos, vezano}`
  - `flags`: `{hasDependents, needsDialog, canPonistenje}`
- Reuse: `modStornoFlow` (`ScanOtpremnica/Zbirna/Prijemnica/Revers`, `GetStornoChainRows`, `GetStornoBlockRows`, `GetChainFlags`), `modPaletniList` čitači.
- Napomena (predikcija, ne mutacija): palete `deltaByMode` = **procena** — za DUPLI/PONIŠTENJE zbir stavki prijemnice po paleti; za ISPRAVKA razlika ako je nova količina poznata (inače 0 dok se ne unese nova).

**W1.2 — Palete-impact čitač**
- Fajl: `modPaletniList.bas` (novi javni čitač, npr. `GetPaleteImpactForPrijemnica(broj) As Collection`).
- Reuse: `GetPaletaAggregates`, detekcija stavki kao u `DetachOsirocenePaletaStavke_TX` (isti ključ `COL_PALS_BROJ_PRIJ` + `Stornirano<>Da`).
- Gotovo: vraća per-paletu `used/cap/neto/amb/preradjena` + delta.

**W1.3 — Test hook**
- Fajl: `modTestStorno.bas` — `Sub Test_BuildStornoImpact()` koji ispiše model za par test dokumenata (Debug.Print).
- Gotovo: model potpun za sve tipove; brojke se poklapaju sa realnim tabelama.
- Rizik: nizak (čist read). Test: `Alt+F8 → Test_BuildStornoImpact`.

---

### FAZA 2 — Panel: Nađi → Uvid → Razlog
*Cilj: browse umesto kucanja; pun brojčani uvid; izbor moda sa pregledom efekta. Zamena `cmb/txt` + `PromptCorrectionMode` MsgBox lanca.*

**W2.1 — Browse podaci**
- Fajl: `modStornoFlow.bas` — `Public Function GetActiveDocumentsForStorno(tipFilter, textFilter) As Collection` (red: `{tip, broj, datum, partner, kolicina, depLevel}`).
- Reuse: `GetTableData` + `Scan*` za `depLevel`.

**W2.2 — „Nađi" overlay** (frmDokumenta, runtime)
- Kontrole: `WithEvents` search TextBox, filter dugmad (Otkup/Otpremnica/Zbirna/Prijemnica), rezultat ListBox. Klik reda → `OpenImpact(tip, broj)`.
- Reuse UI obrazac: `EnsureStorniraniPanel`/`HideBehindPanel`.

**W2.3 — „Uvid" panel**
- Kontrole: kontekst-naslov, ListBox „lanac" (obojeno po `kind`), ListBox „palete" (broj/kapacitet/kg/amb), ListBox „blokovi" (multiselect — preseliti iz postojećeg panela).
- Puni se iz `BuildStornoImpact` (W1.1).

**W2.4 — „Razlog" (mod kartice)**
- 4 dugmeta sa ljudskim jezikom; izbor prebojava lanac i palete brojkama (`deltaByMode`), otključava „Primeni".
- Reuse: `DispatchCorrection` iz postojećeg koda.
- Gotovo: ceo izbor moda bez MsgBox-a; brojčani efekat vidljiv pre Primeni.
- Rizik: srednji (najviše runtime kontrola). Test: Excel po tipu; provera da brojke prate izbor moda.

---

### FAZA 3 — Inline pod-odluke (kraj MsgBox lancima)

**W3.1 — Prelij/preko palete u panelu**
- Fajl: `frmDokumenta.frm` — novi inline korak koji poziva `AdjustPaletaGajbiceZaPrijemnicu_TX(broj, "PRELIJ"/"PREKO")` direktno, umesto `PaletaAdjustPrompt` MsgBox-a.
- Prikaz: kapacitet vizuelno („Paleta P-102: 24/20"), tri dugmeta [Prelij na sledeću] [Slaži preko] [Kasnije].
- Reuse: motor `AdjustPaletaGajbiceZaPrijemnicu_TX` netaknut (samo `spillMode` argument).
- Gotovo: kada višak ne staje → inline izbor u panelu, ne MsgBox.

**W3.2 — „Ne diraj palete"** (DUPLI/PONIŠTENJE) — vraća staru `OTKAZI` mogućnost: storno bez `Detach` → stavke ostaju osirocene (→ Nedovršeno).

**W3.3 — PONIŠTENJE potvrda inline** — zamena `res("blocked")` MsgBox-a punim spiskom posledica u panelu + dugme „Poništi sve".

**W3.4 — Blokovi multiselect** — preseliti postojeći iz „Storno / potvrda" u novi layout (bez logičkih izmena).
- Rizik: srednji. Test: Excel — paletizovana prijemnica sa većom ambalažom (prelij i preko); DUPLI sa „ne diraj palete".

---

### FAZA 4 — ISPRAVKA od početka do kraja
*Cilj: tok koji danas puca na 3 mesta (panel → forma → MsgBox) zatvoriti u jedan.*

**W4.1 — Prijemnica crash-safe**
- Fajl: `frmDokumenta.frm` — u `btnUnosPrij` save putanju uvezati `TryAutoCompleteIspravka(FLOW_DOC_PRIJEMNICA, ...)` (kao otpremnica/zbirna), + `CompletePrijemnicaIspravka` u `modStornoFlow`.
- Reuse: `ReassignPaleteToPrijemnica_TX`, `modStornoContext`.
- Gotovo: ISPRAVKA prijemnice preživi zatvaranje forme (context PENDING → auto-dovrši pri snimanju nove).

**W4.2 — Povratak u panel + prelij/preko inline**
- Posle snimanja nove: umesto `PaletaAdjustPrompt` MsgBox-a → panel prikaže W3.1 korak (kapacitet), pa Rezultat + `CompleteCorrectionContext`.

**W4.3 — Baner „ISPRAVKA u toku"** dok operater unosi novi dokument (vidljiv trag, ne module-level tišina).
- Rizik: srednji-visok (dodiruje save putanju prijemnice). Test: Excel — ISPRAVKA sa promenom količine → prelij/preko; zatvaranje forme usred ISPRAVKE pa nastavak.

---

### FAZA 5 — Nedovršeno (jedan izvor istine) + „Vrati storno"

**W5.1 — Ujedinjeni čitač**
- Fajl: `modStornoFlow.bas` ili novi `modStornoRecovery.bas` — `GetNedovrseno() As Collection`, dedup: `GetPendingCorrections` (tblStornoVeze) + osirocene prijemnice (`GetOsirocenePrijemnice`) + osirocene palete (`GetPrijemniceSaOsirocenimPaletama`).

**W5.2 — Panel „Nedovršeno"** sa akcijom po redu: Prevezi / Skini / Završi ispravku / Odbaci.
- Reuse: `ReassignPrijemnicaToZbirna_TX`, `DetachOsirocenePaletaStavke_TX`, `CompleteCorrectionContext`.

**W5.3 — „Vrati storno"**
- Fajl: `modStorno.bas` (ili `modStornoFlow`) — `UndoStorno_TX(docType, id) As Boolean`: obrni `Stornirano`, reaktiviraj vezane redove, uz iste guard/TX obrasce.
- Odluka (v. §4): ceo lanac ili samo pojedinačni dokument bez zavisnosti.
- Rizik: srednji-visok (reverzija je osetljiva). Test: Excel — storno pa Vrati; provera integriteta.

---

### FAZA 6 — Puna konzistencija + otkupni listovi

**W6.1 — Otkup/Faktura/Novac u framework**
- Fajl: `modStornoFlow.bas` — `ScanOtkup/RunOtkupCorrection`, isto za Faktura/Novac (Faktura PONIŠTENJE povlači novac). `ComboToDocType`/`DispatchCorrection` grane.
- Gotovo: svi tipovi kroz isti panel/model.

**W6.2 — Otkupni list pri stornu bloka**
- Fajl: `modPrint.bas` reuse (`OutputGrupniOtkupniList` / otkup print) — kada se blok stornira, otkupni list označiti/preštampati „STORNIRANO".
- Odluka (v. §4): samo označiti vs auto-reprint.
- Rizik: visok (dodiruje print + širi obuhvat). Test: Excel — storno bloka pa provera otkupnog lista.

---

## 3. Redosled izgradnje i zavisnosti

```
0 (temelj)
└─ 1 (impact motor)  ────────────┐
   └─ 2 (Nađi/Uvid/Razlog) ──┐   │
      └─ 3 (inline odluke) ──┤   │
         └─ 4 (ISPRAVKA e2e) ┘   │
5 (Nedovršeno + Vrati storno) ───┘   (zavisi od 1)
6 (konzistencija + listovi)  ─── zaseban dogovor
```

Preporuka: **0 → 1 → 2 → 3 → 4** daje kompletan objedinjeni tok za glavne dokumente (otpremnica/zbirna/prijemnica + palete + blokovi). Faze 5 i 6 su širenje — pokrenuti tek po potvrdi 0–4.

---

## 4. Otvorene odluke (blokiraju konkretne work-items)

1. **ISPRAVKA unos (W4):** vođeni handoff na postojeću formu (brže, manje rizika) **vs** mini-forma unutar panela (potpuno objedinjeno, veći build). *Predlog: handoff sada, mini-forma kao cilj.*
2. **„Vrati storno" (W5.3):** ceo lanac **vs** samo pojedinačni dokument bez zavisnosti. *Predlog: prvo pojedinačni bez zavisnosti.*
3. **Otkupni list (W6.2):** samo označiti storniran **vs** auto-reprint „STORNIRANO". *Predlog: prvo označiti, reprint opciono.*
4. **Forma:** ostati na runtime overlay-ima u `frmDokumenta` **vs** novi `frmStornoCentar` (traži jednokratni shell u dizajneru). *Predlog: overlay sada.*

---

## 5. Test strategija

- **Statički (svaki commit):** balans `Sub/Function`/`End`, nema duplih `Public` definicija, `file` = ASCII na VBA izvorima, grep ne-ASCII prazno, `git merge-tree` pre integracije `main`-a.
- **Motor bez UI-a (Faza 1):** `modTestStorno` rutine (Debug.Print modela).
- **Excel smoke po fazi:** numerisana checklist (klik po klik) fokusirana na dodato — obavezno: paletizovana prijemnica (prelij/preko), DUPLI/PONIŠTENJE brojke, ISPRAVKA e2e, Vrati storno, Nedovršeno.
- **Regres-čvorovi:** palete motor netaknut (isti pozivi); ISPRAVKA relink identičan `main`-u dok se W4.1 ne uradi.

## 6. Rizici i mitigacije

| Rizik | Mitigacija |
|---|---|
| Runtime kontrole (mnogo, .frx netaknut) | Preslikati proveren obrazac postojećih panela; graditi panel po panel. |
| Dodir save-putanje prijemnice (W4) | Iza flag-a; zadržati postojeći `m_pendingRelinkOldPrij` put dok novi ne prođe smoke. |
| Reverzija storna (W5.3) | Uzak obuhvat prvo (bez zavisnosti); pun TX + guard; integritet provera posle. |
| Print/otkupni listovi (W6) | Posebna faza; ne blokira 0–5. |
| Nemogućnost kompajla u CI | Statička verifikacija + operaterov Excel smoke po fazi. |

---

_Plan v1 · sve posledice (dokumenti · otkupni listovi · palete + kapacitet · ambalaza saldo · faktura · novac) objedinjene u jedan tok. Motori se reuse-uju; menja se orkestracija i uvid._
