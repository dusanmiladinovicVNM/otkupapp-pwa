# Domen — šta dokumenti jesu, nezavisno od koda

> Svrha: da se domen ne rekonstruiše iz poziva. Agent (i čovek) koji prvi put
> dira lanac dokumenata treba da zna šta je invarijanta, a šta samo trenutna
> implementacija — inače „popravi" simptom i razbije pravilo.
>
> Ovo **nije** još jedan opis koda. Gde autoritativni dokument već postoji,
> ovde stoji samo pokazivač.

## 1) Lanac dokumenata

```
Otkupni blok  ->  Otpremnica  ->  Zbirna  ->  Prijemnica  ->  Faktura
 (tblOtkup)     (tblOtpremnica) (tblZbirna) (tblPrijemnica)  (tblFakture
                                                              + tblFakturaStavke)
```

- **Otkupni blok** — jedan otkup od jednog kooperanta, na jednom otkupnom mestu,
  jednog dana. Nosi `BrojDokumenta`, veže se na otpremnicu preko `OtpremnicaID` /
  `BrojOtpremnice`, i nasleđuje `BrojZbirne`.
- **Otpremnica** — roba koja fizički ide sa otkupnog mesta. Više blokova → jedna
  otpremnica.
- **Zbirna** — agregat više otpremnica koje idu istom kupcu/hladnjači.
- **Prijemnica** — prijem robe na odredištu.
- **Faktura** — obračun prema kupcu.

Uporedo, ne u lancu: **`tblAmbalaza`** (ledger kretanja gajbi), **`tblNovac`**
(uplate/isplate), **`tblPaleta` / `tblPaletaStavka`** (paletizacija).

## 2) Invarijante koje kod stvarno drži

Ove su kodirane, ne dogovorene usmeno — izvor je naveden uz svaku.

**Zbirna je agregat, otpremnice su izvor istine.**
`modDokumentInvariant`:

> ZBIRNA = tačno zbir svih svojih AKTIVNIH otpremnica (po `BrojZbirne`).
> KG se proverava **po klasi** (I/II) → hard. Ambalaža ukupno → hard, po klasi →
> soft. Ako se menja otpremnica, mora se validirati i/ili rekalkulisati zbirna.

Praktična posledica: ne „popravljaj" zbirnu upisom u nju. Popravlja se otpremnica,
pa `RecalculateZbirnaFromOtpremnice_TX`.

**Storno nije brisanje.** Dokument-tabele imaju `Stornirano` kolonu; storniran red
ostaje u tabeli i izlazi iz svih agregata. Zato „aktivan" nije isto što i
„postoji". Kaskade (šta storno jednog dokumenta povlači nizvodno) su opisane u
`docs/STORNO_BACKLOG.md` i release notes za `v2.4.0`.

**Koje tabele nose storno je DEKLARISANO, ne pogađa se.** `modSchemaGuard` drži dva
spiska: `STORNO_TABELE` (moraju imati kolonu) i `BEZ_STORNA` (matični podaci, koji
je nemaju). `ExcludeStornirano` je do `v2.78.0` na nenađenu kolonu tiho vraćao
**nefiltrirane** podatke — pa je storniran dokument izlazio kao živ, iz 183 poziva.
Sada se za tabelu iz prvog spiska pada glasno, a `vba_check` pravilom
`STORNO_REGISTAR` ne pušta poziv nad tabelom koju registar ne poznaje.

**Ambalaža je ledger, ne saldo-polje.** `tblAmbalaza` čuva kretanja
(`Smer` = Ulaz/Izlaz, `EntitetID`/`EntitetTip`), a saldo se **izvodi pri čitanju**.
Ne dodavati kolonu sa saldom. Puni model: `docs/AMBALAZA_MODEL.md`.

**Kontekst otpremnice preživljava snimanje otkupnog bloka.** Datum i broj zbirne
ostaju u formi posle snimanja (sledeći blok ide u niz iste otpremnice), kooperant
se briše. Ugovor i testovi: `.claude/rules/otkup-i-dokumenta.md`.

**Registar u `modSchema` je izvor istine za šemu — ne sveska.** Obrnuto je važilo
do PR1: spiskovi kolona osnovnih tabela živeli su isključivo u `.xlsm`, pa se
prazna sveska nije mogla rekonstruisati, a obrisana kolona se videla tek kao pad
upisa satima kasnije.

Sada: `modSchema` deklariše svih 41 tabelu i 590 kolona, `EnsureAllTables` ih
pravi i dopunjava, `VerifySchema` prijavljuje odstupanje (i vrti se u health
check-u), a `SchemaReadyOrFail` je tvrda kapija pred upis. Registar je generisan
iz stvarne sveske i regeneriše se:

```
python tools/dump_schema.py <sveska> --json <put.json>
python tools/gen_schema_module.py --json <put.json>
```

Statička kapija `SEMA_REGISTAR` (`vba_check`) ne pušta `TBL_*` konstantu koje
nema u registru. Suprotan smer — tabela u svesci bez konstante — hvata generator.

**Vlasništvo nad upisom je deklarisano** (`WRITE_OWNERSHIP.json`, ugovor A11).
`python tools/who_writes.py --check-ownership` obara CI na svakog novog pisca
domen-tabele. Lista je zamrznuto zatečeno stanje — račna, ne cilj; skraćuje se
kroz PR-ove.

## 3) Ko šta piše

`WHO_WRITES.md` u ovom folderu — generisana mapa vlasništva nad tabelama
(`python3 tools/who_writes.py --out docs/DOMEN/WHO_WRITES.md`).

Koristi je pre nego što promeniš pravilo upisa: `tblOtkup` piše **12**
produkcionih modula, `tblFakture` **9**. Kad isto polje piše više mesta po
različitim pravilima, to je klasa buga koju test hvata tek posle nastanka.

## 4) Gde je šta autoritativno

| Tema | Autoritet |
|---|---|
| Arhitektura, moduli, tokovi | `docs/ARCHITECTURE_REFERENCE.md`, `docs/ARCHITECTURE_CHANGELOG.md` |
| Ambalaža (ledger, saldo, revers) | `docs/AMBALAZA_MODEL.md` |
| Funkcionalna mapa ekrana | `docs/AgriX_Functional_Map_v142.md` |
| Storno i kaskade | `docs/STORNO_BACKLOG.md`, `docs/STORNO_CENTAR_PLAN_RADA.md` |
| Prerada 2.0 — proizvodno jezgro (model, faze, odluke) | `docs/PRERADA_2_MODEL_I_PLAN.md` |
| SEF (e-fakture) | `docs/SEF_LIFECYCLE_MANUAL.md` |
| Provere integriteta | `docs/INTEGRITET_PROVERE.md` |
| Arhitektonski ugovor (A1–A12, kapije, Pre-Flight) | `docs/DOMEN/ARCHITECTURE_CONTRACT.md` |
| Ciljni model dokumenata (header + stavke, PK/FK, kardinaliteti) | `docs/DOMEN/DOCUMENT_HEADER_LINES.md` |
| Vlasništvo nad upisom (A11) | `docs/DOMEN/WRITE_OWNERSHIP.json` |
| Plan refaktora, redosled PR-ova, kapija odluke | `docs/REFAKTOR_DOKUMENT_HEADER_STAVKE.md` |
| Identitet zbirne, vezivanje prijemnice (ZBR-IDENT-01) | `docs/DOMEN/ZBR_IDENTITET.md` — **superseded posle refaktora** |
| Novac pri stornu / ispravci otkupa | `docs/DOMEN/ODLUKA_NOVAC_PRI_STORNU.md` — **OTVORENO, čeka operatera** |
| Poznata ograničenja | `docs/KNOWN_ISSUES.md` |
| Otkup / dokumenta — pravila izmene | `.claude/rules/otkup-i-dokumenta.md` |
| Verifikacija i definicija gotovog | `CLAUDE.md` §5, `.claude/rules/testovi.md` |

## 5) Šta ovaj folder namerno NE radi

Ne duplira postojeće dokumente i ne opisuje implementaciju. Ako se nešto ovde
razilazi sa kodom, kod je u pravu i **ovaj fajl je bug** — prijavi ga kao i svaki
drugi.
