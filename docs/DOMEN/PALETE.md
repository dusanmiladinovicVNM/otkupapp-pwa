# Palete i prerada — identitet, količina i odluke S8

> **Izvor:** pre-flight S8 (plan `docs/REFAKTOR_DOKUMENT_HEADER_STAVKE.md` §14.58,
> `docs/STANJE_REFAKTORA.md` stavka 273). Mereno nad `main` = `a668c5e6` (grana
> `claude/s8-paleta-preflight`); linije koda u ovom fajlu važe za taj commit.
>
> **Odluke `PAL-ODL-xx` su PREDLOZI dok ih operater ne potvrdi.** Invarijante `PAL-INV-xx`
> imenuju šta kod danas drži i šta mu fali; vlasnik posle S8 je naveden uz svaku.

## 1) Šta je već odlučeno (ne otvara se ponovo)

| Pravilo | Gde |
|---|---|
| Grain paletne stavke je **klasa**, tj. linija prijemnice; ključ linije je `PrijemnicaStavkaID`; `PrijemnicaID` na paletnoj stavci znači **dokument** (labela) | `DOCUMENT_HEADER_LINES.md` §3, §4.6; plan §14.55, odluke 2 i 3 (20.10.2026) |
| Paletizacija ide po stavci, u **istoj transakciji** sa prijemnicom; ista linija se ne paletizuje dvaput (kapija po paru `PrijemnicaID` + `PrijemnicaStavkaID`) | `modDokumenta.SavePrijemnicaMulti_TX:6837`; `modPaletniList.EnsurePrijemnicaNotAlreadyPaletized:193` |
| Storno prijemnice skida palete **u jezgru**, u istoj transakciji, po `PrijemnicaID`; roba na prerađenoj paleti obara ceo storno po imenu | plan §14.55 „Korak 4", odluka 3; `modPaletniList.SkiniPaleteZaStornoPrijemnice:2111` |
| ISPRAVKA prijemnice: palete čekaju novu prijemnicu i prevezuju se po ID-u **obe** strane; DUPLI „ne diraj palete" ostavlja siročad Oporavku | `modDokUnos.PreveziPaleteIspravke:1327`; `modScrOporavak.PreveziPalete:387` |
| Skidanje siročadi po broju (`DetachOsirocenePaletaStavke_TX`) je obrisano; zamena je F8 + Oporavak „Prevezi" | plan §14.55, odluke 24.10.2026, t. 4 (`KI-046`) |
| Prerada uzima palete po `PaletaID`; prerađena paleta se ne stornira, ne zatvara i ne koriguje; storno prerade odbija preradu sa aktivno utovarenom robom | `modPaletniList.SavePrerada_TX:2910-2916`; `modStorno.StornoPaleta:1360`, `StornoPrerada:1529` |
| Prerada 2.0 (lager jedinica, šarža) je **zaseban plan i nije implementiran**; S8 ga ne preduhitruje | `docs/PRERADA_2_MODEL_I_PLAN.md` |

## 2) Identitet

**`PAL-IDENT-01` — mutacija paleta bira dokument po `PrijemnicaID`-u, a liniju po
`PrijemnicaStavkaID`-u; `BrojPrijemnice` i `BrojZbirne` su labele.** Broj prijemnice je
jedinstven samo po kupcu (`modStorno:1848-1851`: „dva kupca istog dana mogu imati
`1/ddmmyy`"; fixture `PRIJEMNICA_BROJ = "1/150326"` kod dva kupca), pa izbor po broju
zahvata tuđi dokument. Storno, prevezivanje i Oporavak to već poštuju. Danas je pravilo
prekršeno u korekciji količina na paletama (plan §14.58, N1).

**`PAL-IDENT-02` — paleta je `PaletaID`, prerada je `PreradaID`; `BrojPalete` i
`BrojPrerade` su labele.** Broj se resetuje po godini (`GenerateBrojPalete:87`,
`GenerateBrojPrerade:2814`). Ekran Palete i utovar rade po ID-u. Alt+F8 makroi biraju po
broju i tekućoj godini i uzimaju prvi pogodak (plan §14.58, N10; `PAL-ODL-04`).

## 3) Invarijante

**`PAL-INV-01` — paletna stavka zna tačnu liniju prijemnice.** `PrijemnicaStavkaID` je
neprazan, postoji u `tblPrijemnicaStavke`, pripada `PrijemnicaID`-u iste paletne stavke,
i klasa se slaže sa klasom stavke.

- **Danas:** važi po konstrukciji. Pisac pravi stavku i paletu u istoj transakciji
  (`modDokumenta:6837`), a prevezivanje i korekcija uzimaju cilj iz kanonskog čitača
  (`modPaletniList:1579`, `:2302`). **Nijedna provera to ne meri**, a prazan ključ je
  dozvoljen kao „red starog oblika" (`modPaletniList:132-136`): takvi su svi redovi
  `tblPaletaStavka` u fixture-u (`tools/make_fixture.py:1962`, *synthetic anomaly*) i
  seed-ovi suite-a Palete.
- **Vlasnik posle S8 (rez 2):** pisac paletne stavke odbija po imenu. Meka provera ide u
  integritet, po stavci. Strog čitač je **po dokumentu**, ne globalan (`KI-081`).

**`PAL-INV-02` — paletizacija je potpuna** (zavisi od `PAL-ODL-02`). Aktivne paletne
stavke jedne stavke prijemnice nose celu stavku: zbir gajbi = gajbe stavke, a zbir neta =
`Kolicina` stavke (±0,5 kg, isti prag kao A3).

- **Danas:** pisac paletizuje celu stavku (`PaletizePrijemnica:140` → `SpillGajbice:2557`),
  a korekcija vraća jednakost (`AdjustPaletaGajbiceZaPrijemnicu_TX:2263`). Meka provera A3
  (`modIntegritet:794`) poredi zbir po paru (`PrijemnicaID`, `Klasa`), ne po stavci, i
  gleda samo linije koje imaju bar jednu aktivnu paletnu stavku.
- **Legitimna odstupanja:**
  - stavka bez gajbi;
  - `PALETIRANJE = NO`;
  - siročad posle ISPRAVKA i DUPLI, dok ih Oporavak ne razreši.

  Delimična paleta i stavka na više paleta su normalno stanje (rubna paleta).
- **Kršenje koje danas niko ne imenuje:** storno palete koja nosi celu stavku aktivne
  prijemnice. A3 tu liniju više ne vidi (`PAL-ODL-06`).

## 4) Odluke S8 — PREDLOZI, čekaju operatera

| ID | Pitanje | Preporuka | Šta zavisi od odgovora |
|---|---|---|---|
| **`PAL-ODL-01`** | Premisa: da li paletizaciju i preradu danas stvarno koristite, kod kog klijenta i koliko često? | — (pitanje, ne predlog) | koliko rigoroza ide u S8: ručna provera štampe i PDF-a, i da li `PAL-ODL-07` uopšte ima korisnika |
| **`PAL-ODL-02`** | Količina: da li je paletizacija **potpuna** (`PAL-INV-02`, jednakost), ili deo stavke sme da ostane van paleta? | **jednakost**, jer je to danas konstrukcija pisca; A3 prelazi na ključ stavke | pravilo pisca i A3 u rezu 2; `PAL-ODL-06` |
| **`PAL-ODL-03`** | Kolone paletne stavke: `BrojPrijemnice`, `BrojZbirne`, `ZbirnaRoditeljID` | **brišu se sve tri**; `PrijemnicaID` ostaje labela dokumenta (kao na liniji fakture, §4.5); `PrijemnicaStavkaID` je jedina veza i obavezan je. `ZbirnaRoditeljID` se izvodi kroz prijemnicu, a dok postoji, drži ga skriven drugi pisac (`ReassignPrijemnicaToZbirna_TX`) | obim reza 2 |
| **`PAL-ODL-04`** | Alt+F8 makroi koji menjaju palete i preradu po broju i tekućoj godini (`ClosePaleta_Prompt`, `StornoPaleta_Prompt`, `SavePrerada_Prompt`, `StornoPrerada_Prompt`, `PaletaAdjust_Prompt`) | četiri koja ekran Palete pokriva po ID-u se **brišu**. `PaletaAdjust_Prompt` ostaje, jer je on put „uradi kasnije": broj se razrešava u **tačno jedan** aktivan dokument, a dvosmislen broj se odbija i imenuje oba kupca | rez 1 |
| **`PAL-ODL-05`** | Siročad klase koju nova prijemnica nema: poruke i `MAPA` B-030 obećavaju „Skini stavke", a ta radnja ne postoji od S6 koraka 7 | **A:** radnja „Skini palete" u Oporavku nad izabranim redom (po `PrijemnicaID` stornirane prijemnice), kroz isto telo kao storno (`SkiniPaletneStavke`), sopstvena transakcija; prerađena paleta se odbija. **B:** samo ispraviti tekst; izlaz ostaje „Storniraj paletu", koji skida i su-stanare | rez 1 (A) ili tekst (B) |
| **`PAL-ODL-06`** | Storno palete koja nosi robu **aktivne** prijemnice (danas se skida sve, i su-stanari) | **odbiti** po imenu dok paleta nosi robu aktivne prijemnice; put je storno ili ispravka prijemnice, odnosno Oporavak. Dozvoljeno ostaje nad paletom bez aktivne robe | rez 2; doslednost sa `PAL-ODL-02` |
| **`PAL-ODL-07`** | Sledljivost na paleti: C-040 (ekran SLEDLJIVOST, obrisan u S1b-3b) i spisak kooperanata na paletnom i preradnom listu (prazan od S1b-1, C-050/C-051). Paleta nosi deo stavke, a stavka je zbir više otkupa, pa je paleta → kooperanti N:M bez alokacije: tačno se može reći samo „kooperanti od kojih je stavka sastavljena" (nadskup) | **prebaciti u S9**: isti graf sledljivosti, S8 daje tačan `PrijemnicaStavkaID`. Paletni i preradni list do S9 izlaze sa zaglavljem i bez spiska, kao danas | da li S8 ima rez 3 |
