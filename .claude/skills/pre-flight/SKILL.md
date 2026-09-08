---
name: pre-flight
description: Kapija PRE pisanja koda za veću izmenu u AgriX/OtkupApp. Koristi kad promena uvodi novu tabelu; domain-relevantnu novu kolonu (identitet/FK, lifecycle/status, količina, datum poslovnog događaja, storno, finansije, sledljivost); novog writer-a nad postojećom tabelom; storno ili kaskadu; promenu identiteta, ključa ili resolver-a; migraciju ili uklanjanje legacy capability-ja; schema/setup; self-update ili import motor; fakturu/SEF/banku; ili tok koji menja kardinalnost, vlasništvo ili lifecycle. NE koristi se za lokalni bug sa jasnom reprodukcijom bez promene contracta, ASCII/Poruka/copy izmene, refaktor jednog modula bez promene read/write semantike, ni čisto vizuelnu UI izmenu bez promene ponašanja.
---

# AgriX Pre-Flight

> Kapija pre koda: postoji da se domen, identitet, nizvodne posledice i dokaz ne otkriju tek u
> review-u. Postojeće politike se ovde **ne prepisuju** — pokazuje se na njih.

## 1. Četiri stvari koje se čitaju, ne pamte

| Šta | Izvor istine |
|---|---|
| DOMAIN — šta dokumenti jesu, lanac, kardinalnost | `docs/DOMEN/README.md` |
| INVARIANTS + ko ih drži | `docs/DOMEN/`, `WHO_WRITES.md`, `modSchemaGuard` (STORNO_REGISTAR), `modDokumentInvariant`, `modIntegritet` |
| DOWNSTREAM | `WHO_WRITES.md` + `docs/RELEASE_GATES.md` §1.2–1.6 |
| PROOF | `.claude/rules/testovi.md`, `CLAUDE.md` §5 |

- Invarijanta **bez poznatog vlasnika** (writer · validator · integrity checker) nije definisana, nego želja.
- DOWNSTREAM pitanje nije „ko zove ovu funkciju" nego **„ko kasnije zavisi od činjenice koju ona upisuje"**.
- Ako odgovor stoji na „uglavnom je 1:1", proveri da li redak N:M slučaj menja arhitekturu.

## 2. BUSINESS EVENTS

Razdvoji najmanje tri, i za svaki reci koji datum, koji status, koja tabela i šta može da postoji bez sledećeg koraka:

- **fizički** — roba stvarno menja stanje ili mesto;
- **poslovni/pravni** — nastaje dokument ili obaveza;
- **finansijski** — nastaje potraživanje/dug ili se novac knjiži.

Ako dva dele isti datum ili status **samo zato što je tako lakše implementirati** → `DOMAIN GAP`, dok se
to ne potvrdi kao poslovno pravilo. `prodaja = utovar = faktura = isporuka` se ne pretpostavlja.

## 3. IDENTITY contract

- Display broj, poslovni broj, naziv, datum — ni bilo koja korisniku vidljiva vrednost — **nisu mutation identity**.
- Canonical ID mora da **putuje** od izabranog reda/objekta do writer-a koji menja podatke.
- `display vrednost → ponovni lookup → canonical ID` je **rizik** kad je ID mogao biti prenet direktno. Prijavi ga, ne ćuti.
- Presedan: `docs/DOMEN/ZBR_IDENTITET.md` (`GeneracijaID` je identitet, `BrojZbirne` je labela).

## 4. CAPABILITY MAP

Namerno se ne zove „parity" — to ime je zauzeto (`tools/vba_parity_check.py`, PWA/VBA parity).
Svaka stara sposobnost završi kao `MIGRATED` / `REPLACED` / `INTENTIONALLY REMOVED`. Sposobnost je više
od dugmeta: unos, pregled, mutacija, validacija, status, upozorenje, oporavak, štampa/PDF/izvoz, dozvole,
pasivni integritetni signal. Katalog je `docs/UI_MIGRACIJA_KATALOG.md`; legacy se ne briše dok red nije
zatvoren. „Kod još postoji" nije dokaz da operater ima funkciju.

## 5. PLATFORM EXPERIMENT

`business ambiguity → specification first` · `platform ambiguity → experiment first`

Za Excel/VBE/COM/MSForms/filesystem ponašanje ne piši arhitekturu na pretpostavci — traži **minimalnu
sondu na Windows + pravom Excelu**. Web sesija je ne može izvršiti (`run_vba` traži Windows + Excel +
`pywin32`), pa se do izvršenja zaključak piše kao `UNMEASURED / NEVERIFIKOVANO`, nikad kao činjenica.

## 6. LANDING

Serijski, ne paralelno: `.claude/`, test harness, CI, schema/setup, centralni UI primitivi, centralni
domen registri. Pre početka: da li grana zavisi od još nemergovane grane, da li druga sesija dira isti
shared fajl, da li se isti test/registry numbering menja paralelno. (`CLAUDE.md` §6, `docs/REFAKTOR_PLAYBOOK.md` §2)

## 7. Izlaz — greppable trag, ne rečenica

`DOMAIN: CLOSED` u chatu nije rezultat. Gde je relevantno, ostaje trag koji se nalazi grep-om:

- imenovan invariant/decision ID u `docs/DOMEN/` (obrazac: `ZBR-IDENT-01`, `ZBR-PARENT-01`);
- red u `docs/UI_MIGRACIJA_KATALOG.md` kad se dira capability;
- `KI-` / `AUD-` / `MIG-` ID **samo kad stvarno postoji otvoren problem**.

Ne upisuj stavku u `KNOWN_ISSUES.md` zato što je pre-flight pokrenut — registar prati probleme, ne procedure.

## 8. Review loop breaker

Kad dva uzastopna review kruga otkriju **novu klasu** problema, a ne još jedan bug iste klase →
**STOP PATCHING**. Imenuj uzrok i vrati se na taj contract pre daljeg patchovanja:

`DOMAIN GAP` · `EVENT GAP` · `IDENTITY GAP` · `INVARIANT GAP` · `DOWNSTREAM GAP` · `CAPABILITY GAP` ·
`PROOF GAP` · `RESPONSIBILITY GAP` · `PLATFORM UNKNOWN` · `LANDING RISK`

## 9. Signali koji se izgovaraju naglas

⚠ DOMAIN NOT CLOSED · ⚠ ARCHITECTURAL EDGE CASE (redak slučaj menja kardinalnost, vlasništvo ili
lifecycle) · ⚠ IDENTITY RISK · ⚠ DOWNSTREAM RISK · ⚠ CAPABILITY RISK · ⚠ FALSE-GREEN RISK (zeleno bez
dokaza da je baš ta tvrdnja merena) · ⚠ PLATFORM UNKNOWN · ⚠ REVIEW TREADMILL

Nezatvorena stavka ne blokira mali lokalni eksperiment — ali eksperiment ne sme neprimetno da postane
produkciona arhitektura.
