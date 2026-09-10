---
paths:
  - "schema/schema.json"
  - "docs/DOMEN/**"
  - "src-vba/modDokumenta.bas"
  - "src-vba/modDokUnos.bas"
  - "src-vba/modDokumentInvariant.bas"
  - "src-vba/modOtkup.bas"
  - "src-vba/modOtkupBlok.bas"
  - "src-vba/modStorno*.bas"
  - "src-vba/modMasterSync.bas"
  - "src-vba/modIntegritet.bas"
  - "src-vba/modSledljivost.bas"
  - "src-vba/modDataAccess.bas"
  - "src-vba/modSetup.bas"
---

# Model pre zakrpe

> Postoji zato što je jedan nedostatak u modelu podataka — dokument nema
> zaglavlje — devet PR-ova plaćan kodom, dok nijedna kapija nije postavila
> pitanje modela. Arheologija sa brojevima:
> `docs/engineering/postmortems/2026-09-model-vs-zakrpa.md`.
> Registar: `docs/DOMEN/KOMPENZACIJE.md`.

## 1) Zašto ovo stoji pored `pre-flight`, a ne u njemu

`pre-flight` dokazuje da je izmena **ispravna unutar zatečenog modela**:
`DOMAIN` / `IDENTITY` / `CARDINALITY` se mere naspram `docs/DOMEN/README.md`.
`GeneracijaID` prolazi svaku od tih osa kao `PROVEN` — ima ugovor
(`ZBR-IDENT-01`), ima ID, ima testove. Kapija je kompenzaciju **overila**.

Dve rupe koje to ostavlja, obe potvrđene:

- **Nijedna osa ne pita da li bi bila prazna da je šema drugačija.** Ose se
  popunjavaju, ne preispituju.
- **`pre-flight` se ne učita za „lokalni bug".** ZBR posao je počeo kao jednoredni
  fix (`ExcludeStornirano`, MIG-005a) — tačno kroz izuzetak iz opisa skilla.
  Zato je ovo `paths:` pravilo, ne skill: okida na fajl, ne na procenu obima.

`pre-flight` §10 (review loop breaker) je bio najbliži i **promašio je smer**:
kaže „imenuj uzrok i vrati se na taj contract" — a contract je bio
`ZBR-IDENT-01`, sama kompenzacija. Izlaz je vodio nazad u zamku.

## 2) Detektor — brojivo, ne osećaj

Prebroj pre pisanja zakrpe. **Dva ili više pogotka = dužnost iz §3.**

| # | Signal | Kako se meri | Presedan |
|---|---|---|---|
| K1 | **Rekonstrukcija** — entitet se sklapa pri čitanju iz više redova/polja | broj call site-ova koji sklapaju | `Split(x, " + ")` na **9** mesta |
| K2 | **Surogat identiteta** — nov ID/ključ/resolver da identifikuje ono što šema ne identifikuje | postoji li kolona koja to nosi | `GeneracijaID`, `ZbirnaIdent*` |
| K3 | **Kapija umesto strukture** — pravilo drži N pisaca, umesto da je strukturno nemoguće | broj pisaca koji moraju da zapamte | **3** pisca „pečate" svaki red |
| K4 | **Isti invariant, više krugova** — ≥2 PR-a ili ≥3 `fix` commita na jednu invarijantu | `git log --grep=<ID>` | 9 PR-ova, 22 `fix` commita |
| K5 | **Duplirana činjenica** — isti podatak na >1 mesta, kod ih drži u koraku | broj mesta + ko sinhronizuje | broj + vlasnik + generacija kao „scope" |

Signal nije bug. Signal je **da se nedostatak modela plaća kodom** — i da cenu
treba izgovoriti pre nego što se plati još jedna rata.

## 3) Dužnost dva rešenja — uvek obe kolone, pre koda

Kad detektor okine, u chat ide ova tabela **pre** implementacije:

| | ZAKRPA | MODEL |
|---|---|---|
| Šta se menja | | |
| Dodirnuti fajlovi / pisci / call site-ovi | | |
| **Šta NESTAJE** | (obično ništa) | |
| Cena sada | | |
| Cena ako se ne uradi (sledeće rate) | | |

Zatim jedna rečenica preporuke i **razlog**.

Zakrpa sme da pobedi. Ne sme da pobedi **prećutno** — ovo je pravilo o tome da
poređenje postoji, ne o tome ko pobeđuje.

## 4) „Šta nestaje" je broj koji odlučuje

To je jedina kolona koja meri **širu sliku**, i jedina koja se nikad nije
pojavila tokom devet PR-ova. Za ZBR posao njena vrednost je bila: `GeneracijaID`
(355 pojava u `src-vba/`), `ZbirnaIdent*` (74), `VlasniciPoBroju` (43),
`RequireJedanVlasnik*` (20), 9 `Split` mesta, ceo ugovor `ZBR-IDENT-01` i njegove
kapije. `DOCUMENT_HEADER_LINES.md` §10 to danas kaže otvoreno: sa headerom taj
identitet **postoji direktno**, pa resolveri i vlasničke kapije „nemaju posao".

Broj se piše kao merenje sa komandom, ne kao procena.

## 5) Cena modela se meri, ne pretpostavlja

Model se odbacuje kao „preskup" bez merenja — a najskuplja stavka mu je obično
migracija. Zato se **pre** odbacivanja proveri i izgovori:

- **Ima li legacy podataka?** Za ovaj refaktor odgovor je bio **ne**
  (`REFAKTOR_DOKUMENT_HEADER_STAVKE.md`, uvod: nema migracije, sezona prošla,
  nijedan klijent na starom programu) — dakle „stari model se briše, ne prevodi".
  Ta činjenica je stajala u repou i **nijedan korak je nije tražio**.
- Je li šema kanon u gitu (`schema/schema.json`), pa je izmena diff a ne migracija?
- Koliko postojećih zakrpa umire sa modelom (§4)?

Bez ova tri odgovora rečenica „model je preskup" je pretpostavka, i tako se
prijavljuje (`CLAUDE.md` §1.4).

## 6) Kad zakrpa pobedi — red u registru, iste sesije

`docs/DOMEN/KOMPENZACIJE.md` dobija/uvećava red: simptom · gde je model kriv ·
izabrana zakrpa · odbačena alternativa modela · cena do sada · broj pogodaka.

Ovo je jedini deo mehanizma koji **preživljava kraj sesije**. Obrazac se ne vidi
u jednom PR-u — vidi se u zbiru, a zbir niko ne pamti. Registar pretvara
„primeti obrazac" (nemoguće) u „pročitaj fajl" (trivijalno).

## 7) Treći pogodak nije opcion

Kad isti red u registru dobije **treći** pogodak, zakrpa prestaje da bude
podrazumevana: ide `pre-flight` verdikt za izmenu modela, ili eksplicitna
korisnikova odluka da se i dalje krpi — zapisana u registru.

## 8) Šta ovo pravilo NIJE

- **Nije licenca za redizajn.** Jedan pogodak je zapažanje, ne akcija.
  `minimal change over idealized redesign` i dalje važi — meri se samo preko svih
  rata koje isti nedostatak traži, ne preko jedne.
- **Ne uvodi slojeve.** Granice iz `DOCUMENT_HEADER_LINES.md` §8 važe:
  bez `tblDokumenti`/EAV, bez Repository/App framework-a, bez `Cmd`/`Qry`/`Rules`,
  bez alokacionih tabela bez poslovnog zahteva.
- **Ne pomera A11 vlasništvo upisa** niti zaobilazi `pre-flight`. Izmena modela
  **pojačava** te kapije — nova tabela/kolona i dalje ide kroz `schema.json`,
  `gen_schema_module.py --check` i `who_writes.py --check-ownership`.
- **Nije retroaktivna revizija.** Ne otvaraj zatvorene ugovore zato što je pravilo
  novo; primenjuje se na posao koji sada počinje.
