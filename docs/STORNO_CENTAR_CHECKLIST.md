# Storno centar — lista za proveru (Faze 0–4)

> Ručni test-plan za granu `claude/storno-dedup-ux-k6esuw`. Prolazi po funkciji;
> svaka stavka: **koraci → očekivano → [ ] ✓/✗**. VBA se ne kompajlira u CI —
> ovo je operaterov smoke-test u Excelu. Na kraju: šta je automatabilno za
> buduće regres-test module.
>
> Pokriveni commit-ovi: prijemnica-u-framework, Faza 0 (popravke), Faza 1
> (agregator), Faza 2a (uvid), 2b (browse), Faza 3 (ne diraj palete), Faza 4
> (ISPRAVKA crash-safe).

## 0. Priprema
- [ ] `git fetch/checkout/pull` grane; u Excelu `Alt+F8 → ImportAllVBA`.
- [ ] `Debug → Compile VBAProject` — **bez ijedne greške** (naročito: nema „Ambiguous name", nema „Sub/Function not defined"). Snimi.
- [ ] Otvori `Ctrl+G` (Immediate) za test hook-ove.

---

## 1. Agregator uvida — `Test_BuildStornoImpact` (Faza 1)
Cilj: model uticaja se slaže sa STVARNIM stanjem tabela (temelj celog panela).

- [ ] `Alt+F8 → Test_BuildStornoImpact` → broj **paletizovane prijemnice**, tip „Prijemnica".
  - Očekivano u Immediate: `Header` (partnerID/datum/kolicina), `Summary` (blocks/palete + detach gajb/neto/amb), red po paleti `used/cap | this=gajb (neto, amb)`.
- [ ] **Poklapanje:** `used/cap` = stvarno stanje palete; `this` (detach) = zbir gajbi/kg/amb TE prijemnice na toj paleti.
- [ ] Ponovi za **Otpremnicu** i **Zbirnu** (palete se vide preko BrojZbirne).
- [ ] Prijemnica **bez paleta** → `palete=0`, prazna palete-sekcija (nema greške).
- [ ] Prerađena paleta u obuhvatu → red ima `[PRERADJENA]`.

---

## 2. Browse „Nađi za storno" (Faza 2b)
- [ ] U storno sekciji postoji dugme **„Nađi za storno"** (ispod „Osiroceni dokumenti"). Ako pozicija preklapa nešto → zabeleži (lako se pomeri).
- [ ] Klik → full-screen lista aktivnih dokumenata (Tip/Broj/Datum/Partner/Količina).
- [ ] **Filter tipa** (combo: Svi/Prijemnica/Otpremnica/Zbirna) → lista se sužava.
- [ ] **Pretraga** (kucaj deo broja) → lista se filtrira uživo; prazan pojam = sve.
- [ ] Broj u naslovu („Nađi… (N)") = broj redova.
- [ ] Izbor reda + **„Otvori storno"** (ili **dvoklik**) → otvara se Storno panel za taj dokument.
- [ ] „Zatvori" vraća na formu bez promena.
- [ ] Distinct po broju: prijemnica sa Klasa I+II se pojavljuje **jednom**.

---

## 3. Panel „Uvid" (Faza 2a)
Otvori storno panel (preko „Nađi" ili stari put: tip + broj + Storno).

- [ ] **Header traka**: partner/stanica, datum, količina.
- [ ] **Lanac (Dotaknuti dokumenti)**: red po dokumentu + „šta se dešava".
- [ ] **Palete lista** (samo ako ima paleta): `Paleta[+PRERAD] | used/cap | skida gajb | kg | amb`.
- [ ] **Summary traka** dole: „dokumenti u lancu / blokovi / palete (DUPLI/PONIŠTENJE skida: gajb/kg/amb)". Brojke = kao `Test_BuildStornoImpact`.
- [ ] **Blokovi lista** (samo ako ima blokova) sa checkbox po redu.
- [ ] **Raspored po prisustvu**: bez paleta → nema palete-liste; bez blokova → nema blok-liste; visina se lepo raspodeli.

---

## 4. Četiri moda — Prijemnica
Za paletizovanu prijemnicu (ima palete + eventualno faktura + blokovi):

- [ ] **DUPLI** → prijemnica stornirana, poruka „paletne stavke skinute: N". Proveri: paleta-`NetoKg`/`AmbalazaKg`/`BrojGajbica` **smanjeni** za `this` iznos; prazna paleta → stornirana; ispod kapaciteta → `Otvorena`.
- [ ] **PONIŠTENJE** → prvo MsgBox pun spisak posledica + potvrda; na DA isto skida palete + oslobađa fakturu. Dugme aktivno samo kad ima zavisnosti.
- [ ] **REŠI KASNIJE** → prijemnica **ostaje aktivna**, samo recovery zapis (Nedovršeno). Palete netaknute.
- [ ] **ISPRAVKA** → stara stornirana, forma prefill-ovana, fokus na količinu (v. sekcija 7).

---

## 5. Multiselect otkupnih blokova (samostalni)
- [ ] Blokovi lista prikazuje BrDok/Količina/Klasa/Kooperant.
- [ ] **Bez čekiranja** → DUPLI/PONIŠTENJE: blokovi ostaju (kod otpremnice/zbirne se oslobađaju; kod prijemnice ostaju vezani za zbirnu).
- [ ] **Čekiran blok** → posle moda poruka „Otkupni blokovi dodatno stornirani: N"; taj blok je `Stornirano=Da`, nečekirani aktivni.
- [ ] Otkupni list čekiranog bloka: (za sada) blok je storniran; auto-reprint dolazi u Fazi 6.

---

## 6. „Ne diraj palete" (Faza 3, W3.2)
Za **prijemnicu sa paletama**:

- [ ] U summary redu (desno) postoji checkbox **„Ne diraj palete"**.
- [ ] Čekiraj → **DUPLI** → poruka „Palete ostavljene osirocene → Osiroceni dokumenti (Mod: Palete)". Paleta **NIJE** dirana (kg/amb nepromenjeni), stavke ostaju.
- [ ] Osirocene stavke se vide u „Osiroceni dokumenti → Mod: Palete".
- [ ] **Bez čekiranja** → DUPLI skida palete (kao sekcija 4).
- [ ] Za otpremnicu/zbirnu/dokument bez paleta → checkbox se **ne prikazuje**.

---

## 7. ISPRAVKA — od početka do kraja + crash-safe (Faza 4, W4.1)
- [ ] **In-session:** Storno prijemnice → ISPRAVKA → forma prefill-ovana (ista roba). Unesi ispravnu → Snimi → palete **prevezane** na novu (poruka), ispravka završena, „Osiroceni" ne prijavljuje.
- [ ] **Promena količine:** unesi drugačiju ambalažu → posle snimanja MsgBox **prelij/preko** (kad višak ne staje: „+N gajb. ne staje, slobodno X od Y"): DA=prelij (dopuni pa višak na sledeću), NE=preko (svesno preko kapaciteta), OTKAZI=kasnije.
- [ ] **Crash-safe:** Storno prijemnice → ISPRAVKA → **ZATVORI formu Dokumenta** (ne unosi novu). Ponovo otvori → unesi novu prijemnicu (istu zbirnu) → Snimi → pitanje „Čeka ISPRAVKA za storniranu X — zamena?" → DA → palete prevezane, ispravka završena.
- [ ] **Safe-stop:** ako ima **više** PENDING ISPRAVKI prijemnice → pri snimanju upozorenje da auto-prevezivanje NIJE urađeno (ne bira naslepo).
- [ ] **NE (odbij):** na pitanje „zamena?" → NE → običan, nepovezan unos (stara ispravka ostaje u Nedovršeno).

---

## 8. Palete — integritet (fokus)
- [ ] DUPLI/PONIŠTENJE: paleta-header = zbir preostalih aktivnih stavki (self-heal), uključuje su-stanare (druge prijemnice na istoj paleti **netaknute**).
- [ ] Prerađena paleta u obuhvatu detach-a → **NIJE** stornirana, ide upozorenje „preradjena".
- [ ] `BrutoKg` = neto + amb + težina palete (preračunat posle svake izmene).
- [ ] Ambalaza-**saldo** (ledger) obrnut na svakom stvarnom stornu prijemnice (proveri karticu ambalaže kooperanta/firme).

---

## 9. Regresija (staro mora i dalje da radi)
- [ ] **Stari ulaz:** tip u `cmbStornoDokument` + kucaj broj + Storno → radi (framework tipovi otvaraju panel; ostali stari put).
- [ ] **Otpremnica / Zbirna** kroz panel: uvid + 4 moda kao pre + palete/summary vidljivi.
- [ ] **Revers** (izdavanje/povrat/OM): kratak DA/NE izbor, nepromenjen.
- [ ] **Otkup / Faktura / Novac**: prost „Stornirati? Da/Ne", nepromenjen (framework ih preuzima u Fazi 6).
- [ ] **Pregled storniranih** panel radi.
- [ ] **Osiroceni dokumenti** (recovery): Prevezi prijemnicu / Skini palete rade.
- [ ] Snimanje **nove** otpremnice/zbirne/prijemnice (bez ispravke) radi normalno; auto-štampa nepromenjena.

---

## 10. Perf (Faza 0, W0.2)
- [ ] Storno panel nad zbirnom sa **puno** otkupnih blokova → lista se puni bez primetnog zastoja.

---

## Poznata ograničenja (svesno, NIJE bug)
- Partner u browse/header je prikazan kao **ID** (rezolucija imena = kasnija polish).
- **Prelij/preko** (W3.1) i **potvrda poništenja** (W3.3) ostaju modalni `MsgBox` (runtime overlay nije modalan; namenska forma = zaseban korak).
- Browse pokriva **Prijemnica/Otpremnica/Zbirna** (Otkup/Faktura/Novac/Revers → Faza 6).
- „Nedovršeno" je još **dva izvora** (recovery panel + tblStornoVeze) → ujedinjenje je Faza 5.
- „Vrati storno" → Faza 5.

---

## Za buduće REGRES-TEST module (automatabilno bez UI-a)
Ove funkcije su čiste/read-only ili TX sa jasnim ulazom → pišu se `modTestStorno` rutine (Debug.Print/asertacije), brane regresiju bez klika:

| Funkcija | Šta testirati |
|---|---|
| `modStornoImpact.BuildStornoImpact` | model = zbir stvarnih tabela (header/chain/palete/summary) |
| `modPaletniList.GetPaleteImpactByField` | per-paleti used/cap + this-detach = stvarno stanje; prerađena flag |
| `modStornoFlow.GetActiveDocumentsForStorno` | distinct po broju, filter tip+tekst, samo aktivni |
| `modStornoFlow.GetStornoChainRows / GetStornoBlockRows` | pravi lanac/blokovi po tipu |
| `modStornoFlow.StornoSelectedBlocks_TX` | storno N blokova atomično; -1 na neispravan ID (rollback) |
| `modStornoFlow.RunPrijemnicaCorrection` (svi modi + `skipPalete`) | DUPLI skida ⇔ skipPalete preskače; PONIŠTENJE gate; ISPRAVKA needsForm+context |
| `modStorno.*_TX` (postojeće) | soft-delete + kaskade + ambalaza ledger (već ima `modTestStorno`/`modBusinessFlowProTests`) |

Predlog obrasca: `Test_<Funkcija>` sub-ovi koji naprave fixture (ili čitaju poznati dokument), pozovu funkciju, i `Debug.Assert`/`Debug.Print` očekivano vs dobijeno. Reci kad hoćeš da ih napišem (Faza „test harness").

---

_Checklist v1 · grana `claude/storno-dedup-ux-k6esuw` · Faze 0–4._
