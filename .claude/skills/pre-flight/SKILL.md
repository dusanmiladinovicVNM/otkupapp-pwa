---
name: pre-flight
description: Kapija PRE pisanja koda za veću izmenu u AgriX/OtkupApp. Koristi kad promena uvodi novu tabelu; domain-relevantnu novu kolonu (identitet/FK, lifecycle/status, količina, datum poslovnog događaja, storno, finansije, sledljivost); novog writer-a nad postojećom tabelom; storno ili kaskadu; promenu identiteta, ključa ili resolver-a; migraciju ili uklanjanje legacy capability-ja; schema/setup; self-update ili import motor; fakturu/SEF/banku; ili tok koji menja kardinalnost, vlasništvo ili lifecycle. NE koristi se za lokalni bug sa jasnom reprodukcijom bez promene contracta, ASCII/Poruka/copy izmene, refaktor jednog modula bez promene read/write semantike, ni čisto vizuelnu UI izmenu bez promene ponašanja — ali ti izuzeci prestaju da važe čim posao tokom rada dodirne dedup/grupisanje/brojanje, parent linkage, mapiranje reda u logički dokument, ownership ili lifecycle (§2).
---

# AgriX Pre-Flight

> Kapija pre koda: postoji da se domen, identitet, nizvodne posledice i dokaz ne otkriju tek u
> review-u. Postojeće politike se ovde **ne prepisuju** — pokazuje se na njih.

## 1. Verdikt pre koda — blokira

Pre produkcionog koda izgovori status za svaku osu, uz **dokaz** (fajl:linija, ID, ime testa)
ili `N/A — razlog`:

`DOMAIN` · `IDENTITY` · `CARDINALITY` · `INVARIANTS/OWNER` · `WRITERS` · `DOWNSTREAM` ·
`CAPABILITY` · `ACCEPTANCE PROOF` · `PLATFORM` · `LANDING` → `PROVEN` / `GAP` / `N/A`

**`GAP` na `DOMAIN`, `IDENTITY`, `CARDINALITY`, `INVARIANTS/OWNER` ili `DOWNSTREAM`, kad je ta osa
relevantna → NEMA PRODUKCIONOG KODA.** Dozvoljen je samo eksperiment koji je tako i označen i ne
ulazi u produkcioni put. `PROVEN` bez navedenog dokaza je `GAP`, ne `PROVEN`.

Nijedan checker ovo ne meri — zato ide dokaz uz svaku osu, a ne sama reč.

## 2. Escalation — bez obzira kako je posao počeo

Trigger u opisu gleda **početak** posla, a posao mutira. MIG-005a je počeo kao jednoredni fix
(`ExcludeStornirano`), pa je §28.1f MIG-005b ostavio otvoren sa netačnom tvrdnjom da picker „nosi
samo broj" — a §28.1g pokazuje da mu je `GeneracijaID` sve vreme bio dostupan.

Zato: čim tokom rada dodirneš **dedup / group / count**, parent linkage, lookup po poslovnom broju,
mapiranje reda u logički dokument, ownership, lifecycle/storno — ili moraš da **promeniš premisu
postojećeg regression testa** — STOP i uradi verdikt iz §1 pre nastavka. Izuzetak „lokalni bug sa
jasnom reprodukcijom" tada više ne važi.

## 3. Četiri stvari koje se čitaju, ne pamte

| Šta | Izvor istine |
|---|---|
| DOMAIN — šta dokumenti jesu, lanac, kardinalnost | `docs/DOMEN/README.md` |
| INVARIANTS + ko ih drži | `docs/DOMEN/`, `WHO_WRITES.md`, `modSchemaGuard` (STORNO_REGISTAR), `modDokumentInvariant`, `modIntegritet` |
| DOWNSTREAM | `WHO_WRITES.md` + `docs/RELEASE_GATES.md` §1.2–1.6 |
| PROOF | `.claude/rules/testovi.md`, `CLAUDE.md` §5 |

- Invarijanta **bez poznatog vlasnika** (writer · validator · integrity checker) nije definisana, nego želja.
- DOWNSTREAM pitanje nije „ko zove ovu funkciju" nego **„ko kasnije zavisi od činjenice koju ona upisuje"**.
- Ako odgovor stoji na „uglavnom je 1:1", proveri da li redak N:M slučaj menja arhitekturu.
- **Uzročna tvrdnja traži creation path.** „X može / ne može nastati" dokazuje se čitanjem
  produkcionog write path-a koji X pravi. Test, fixture, komentar i zatečen podatak nisu dokaz
  poslovnog mehanizma: `ZBR_IDENTITET.md` §11b je tačno opisivao kod a pogrešno mehanizam, sve dok
  import path nije pročitan.

## 4. BUSINESS EVENTS

Kad izmena dodiruje poslovni tok robe, dokumenta ili novca, razdvoji najmanje tri, i za svaki reci
koji datum, koji status, koja tabela i šta može da postoji bez sledećeg koraka:

- **fizički** — roba stvarno menja stanje ili mesto;
- **poslovni/pravni** — nastaje dokument ili obaveza;
- **finansijski** — nastaje potraživanje/dug ili se novac knjiži.

Inače `EVENTS: N/A — razlog` (self-update, schema/setup, čist identity refaktor). Tri reda se ne
popunjavaju ritualno — prazan red obara signal celog verdikta.

Ako dva dele isti datum ili status **samo zato što je tako lakše implementirati** → `DOMAIN GAP`, dok
se to ne potvrdi kao poslovno pravilo. `prodaja = utovar = faktura = isporuka` se ne pretpostavlja.

## 5. IDENTITY contract

- Display broj, poslovni broj, naziv, datum — ni bilo koja korisniku vidljiva vrednost — **nisu mutation identity**.
- Canonical ID mora da **putuje** od izabranog reda/objekta do writer-a koji menja podatke.
- `display vrednost → ponovni lookup → canonical ID` je **rizik** kad je ID mogao biti prenet direktno. Prijavi ga, ne ćuti.
- **Pre testa koji broji, grupiše ili deduplikuje napiši koji semantički nivo meriš:** fizički red ·
  logički dokument · poslovni broj. Obrazac je A6 u `ZBR_IDENTITET.md` — preduslov „2 fizička reda,
  ista `GeneracijaID`", tvrdnja `activeLogicalCount = 1`.
- **Fixture nosi oznaku** `normal production state` ili `synthetic anomaly / fault injection`. Iz
  anomalnog fixture-a se ne izvodi poslovna invarijanta.
- Presedan: `docs/DOMEN/ZBR_IDENTITET.md` (`GeneracijaID` je identitet, `BrojZbirne` je labela).

## 6. CAPABILITY MAP

Namerno se ne zove „parity" — to ime je zauzeto (`tools/vba_parity_check.py`, PWA/VBA parity).
Svaka stara sposobnost završi kao `MIGRATED` / `REPLACED` / `INTENTIONALLY REMOVED`. Sposobnost je više
od dugmeta: unos, pregled, mutacija, validacija, status, upozorenje, oporavak, štampa/PDF/izvoz, dozvole,
pasivni integritetni signal. Katalog je `docs/UI_MIGRACIJA_KATALOG.md`; legacy se ne briše dok red nije
zatvoren. „Kod još postoji" nije dokaz da operater ima funkciju.

## 7. PLATFORM EXPERIMENT

`business ambiguity → specification first` · `platform ambiguity → experiment first`

Za Excel/VBE/COM/MSForms/filesystem ponašanje ne piši arhitekturu na pretpostavci — traži **minimalnu
sondu na Windows + pravom Excelu**. Web sesija je ne može izvršiti (`run_vba` traži Windows + Excel +
`pywin32`), pa se do izvršenja zaključak piše kao `UNMEASURED / NEVERIFIKOVANO`, nikad kao činjenica.

## 8. LANDING

Serijski, ne paralelno: `.claude/`, test harness, CI, schema/setup, centralni UI primitivi, centralni
domen registri. Pre početka: da li grana zavisi od još nemergovane grane, da li druga sesija dira isti
shared fajl, da li se isti test/registry numbering menja paralelno. (`CLAUDE.md` §6, `docs/REFAKTOR_PLAYBOOK.md` §2)

## 9. Izlaz — greppable trag, ne rečenica

Verdikt iz §1 u chatu je uslov, ali ne i trag. Gde je relevantno ostaje i nešto što se nalazi grep-om:

- imenovan invariant/decision ID u `docs/DOMEN/` (obrazac: `ZBR-IDENT-01`, `ZBR-PARENT-01`);
- red u `docs/UI_MIGRACIJA_KATALOG.md` kad se dira capability;
- `KI-` / `AUD-` / `MIG-` ID **samo kad stvarno postoji otvoren problem**.

Ne upisuj stavku u `KNOWN_ISSUES.md` zato što je pre-flight pokrenut — registar prati probleme, ne procedure.

## 10. Review loop breaker

Kad dva uzastopna review kruga otkriju **novu klasu** problema, a ne još jedan bug iste klase →
**STOP PATCHING**. Imenuj uzrok i vrati se na taj contract pre daljeg patchovanja:

`DOMAIN GAP` · `EVENT GAP` · `IDENTITY GAP` · `INVARIANT GAP` · `DOWNSTREAM GAP` · `CAPABILITY GAP` ·
`PROOF GAP` · `RESPONSIBILITY GAP` · `PLATFORM UNKNOWN` · `LANDING RISK`

## 11. Signali koji se izgovaraju naglas

⚠ DOMAIN NOT CLOSED · ⚠ ARCHITECTURAL EDGE CASE (redak slučaj menja kardinalnost, vlasništvo ili
lifecycle) · ⚠ IDENTITY RISK · ⚠ DOWNSTREAM RISK · ⚠ CAPABILITY RISK · ⚠ FALSE-GREEN RISK (zeleno bez
dokaza da je baš ta tvrdnja merena) · ⚠ PLATFORM UNKNOWN · ⚠ REVIEW TREADMILL

Nezatvorena stavka ne blokira mali lokalni eksperiment — ali eksperiment ne sme neprimetno da postane
produkciona arhitektura.
