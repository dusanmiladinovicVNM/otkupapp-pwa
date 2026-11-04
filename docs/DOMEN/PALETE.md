# Palete i prerada — identitet, količina i odluke S8

> **Izvor:** pre-flight S8 (plan `docs/REFAKTOR_DOKUMENT_HEADER_STAVKE.md` §14.58,
> `docs/STANJE_REFAKTORA.md` stavka 273). Mereno nad `main` = `a668c5e6` (grana
> `claude/s8-paleta-preflight`); linije koda u ovom fajlu važe za taj commit.
>
> **Odluke `PAL-ODL-01..07` je operater doneo 03.11.2026** (§4; `docs/STANJE_REFAKTORA.md` stavka 274).
> Invarijante `PAL-INV-xx` imenuju šta kod danas drži i šta mu fali; vlasnik posle S8 je naveden
> uz svaku.

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
zahvata tuđi dokument. Storno, prevezivanje i Oporavak to već poštuju. Korekcija količina
na paletama ga je kršila (plan §14.58, N1); od reza 1a (stavka 277) bira dokument po
`PrijemnicaID`-u, a paletne stavke po paru ID-eva, a Alt+F8 razrešava broj u tačno jedan
aktivan dokument. Ostaje rezervna grana po broju u prevezivanju (N2, rez 1b).

**`PAL-IDENT-02` — paleta je `PaletaID`, prerada je `PreradaID`; `BrojPalete` i
`BrojPrerade` su labele.** Broj se resetuje po godini (`GenerateBrojPalete:87`,
`GenerateBrojPrerade:2814`). Ekran Palete i utovar rade po ID-u. Alt+F8 makroi biraju po
broju i tekućoj godini i uzimaju prvi pogodak (plan §14.58, N10; `PAL-ODL-04`).

## 3) Invarijante

**`PAL-INV-01` — paletna stavka zna tačnu liniju prijemnice.** `PrijemnicaStavkaID` je
neprazan, postoji u `tblPrijemnicaStavke`, pripada `PrijemnicaID`-u iste paletne stavke,
i klasa se slaže sa klasom stavke.

- **Vlasnik (S8 rez 2b-1, `PAL-ODL-13`):** pisac paletne stavke odbija po imenu, pre ijedne
  mutacije — `modPaletniList.RequireLinijaPalete` na vrhu `SpillGajbice`, jedinog puta do
  `AddStavka` (paletizacija i korekcija). Odbija prazan, nepostojeći, tuđ ključ, ključ druge
  klase i ključ koji nose dve stavke. Čita samo redove tog ključa, ne globalni strog čitač
  (`KI-081`).
- **Pravilo je jedno:** `modPaletniList.PalInv01Razlog` daje razlog i piscu (tvrdo) i
  integritetu (meko, C1 u `modIntegritet`). Aktivna paletna stavka bez ključa je nalaz C1; A3
  je ne meri, jer se ne zna čija je linija.
- **Prevezivanje** upisuje ključ iz kanonskog čitača ciljnog dokumenta (`modPaletniList`,
  `ReassignPaleteToPrijemnica_TX`), pa pravilo važi po konstrukciji; kapija tamo nije
  dodata, jer pogrešan ključ tim putem ne može da nastane, a kapija koju nijedan slučaj ne
  doseže ne bi ništa merila.
- **Fixture i seed-ovi** nose ključ linije (`tools/make_fixture.py`, `modTestStorno`). Red bez
  ključa postoji samo kao *synthetic anomaly* u testu (Palete T14, BFP
  `Test_S8R2b1_PisacPaleteTraziLiniju`).

**`PAL-INV-02` — paletizacija je potpuna** (`PAL-ODL-02`, odlučeno 03.11.2026). Aktivne paletne
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
  prijemnice. A3 tu liniju više ne vidi. Po `PAL-ODL-06` storno takve palete se odbija.
  Od S8 reza 1c (`docs/STANJE_REFAKTORA.md` stavka 290) to drži jezgro `modStorno.StornoPaleta`
  kroz `modPaletniList.PaletaAktivnaRobaRazlog`; A3 po stavci ostaje za rez 2a.

## 4) Odluke S8 (operater, 03.11.2026)

| ID | Odluka |
|---|---|
| **`PAL-ODL-01`** | **Palete i prerada će se slabo koristiti, možda ni do sledećeg leta.** Posle ovog refaktora, od januara 2027, ide Prerada 2.0 u detalje. Posledica: S8 ne gradi ništa što Prerada 2.0 zamenjuje i ne širi se van identiteta i veza. Ručna provera štampe i PDF-a nije potrebna dok rez ne menja papir |
| **`PAL-ODL-02`** | **Paletizacija je potpuna** (`PAL-INV-02`): zbir aktivnih paletnih stavki jedne stavke = cela stavka; A3 prelazi na ključ stavke |
| **`PAL-ODL-03`** | **Brišu se sve tri kolone** (`BrojPrijemnice`, `BrojZbirne`, `ZbirnaRoditeljID`). `PrijemnicaStavkaID` je jedina veza i obavezan je; `PrijemnicaID` ostaje labela dokumenta; brojevi za prikaz se čitaju iz zaglavlja prijemnice |
| **`PAL-ODL-04`** | **Četiri makroa se brišu** (`ClosePaleta_Prompt`, `StornoPaleta_Prompt`, `SavePrerada_Prompt`, `StornoPrerada_Prompt`); ekran Palete ih pokriva po ID-u. **`PaletaAdjust_Prompt` ostaje strog:** broj se razrešava u tačno jedan aktivan dokument, a dvosmislen broj se odbija i imenuje oba kupca |
| **`PAL-ODL-05`** | **A: radnja „Skini palete" u Oporavku**, nad izabranim redom (po `PrijemnicaID` stornirane prijemnice), kroz postojeće telo storna (`SkiniPaletneStavke`), u sopstvenoj transakciji; prerađena paleta se odbija. Poruke prevezivanja i `MAPA` B-030 pokazuju na tu radnju. **U istom rezu (1b), pre radnje:** odluka „paleta je prazna" prelazi sa prikaznog čitača na strog brojač aktivnih stavki (N8, `S8-REV-05`) |
| **`PAL-ODL-06`** | **Storno palete se odbija po imenu dok paleta nosi robu aktivne prijemnice**; put je storno ili ispravka prijemnice. Paleta bez aktivne robe (siročad, prazna) i dalje se stornira |
| **`PAL-ODL-07`** | **Sledljivost na paleti ide u S9**: C-040 i spisak kooperanata na paletnom i preradnom listu. S8 obezbeđuje tačan `PrijemnicaStavkaID`; do S9 list izlazi sa zaglavljem, bez spiska |

Odluke za rez 2a (operater, 04.11.2026; plan §14.59):

| ID | Odluka |
|---|---|
| **`PAL-ODL-08`** | **A3 meri paletizovane stavke:** stavku prijemnice koja ima bar jednu paletnu stavku (aktivnu ili storniranu); zbir aktivnih mora biti cela stavka (`PAL-INV-02`). Roba koja nikad nije paletizovana nije nalaz |
| **`PAL-ODL-09`** | **Kolona „Zbirna" u listi stavki palete (C-048)** se čita iz zaglavlja zbirne preko veze prijemnice, ne sa paletne stavke i ne kao nov čitalac broja zbirne na prijemnici |
| **`PAL-ODL-10`** | **`KI-086` i `KI-087` idu u zaseban rez posle 2a, pre 2b** (ispravnost kapija i ishoda, ne identitet) |
| **`PAL-ODL-11`** | **2a ide kao dva PR-a:** 2a-1 storno zbirne (R1–R4, `KI-039`), 2a-2 labele, C-048 i integritet (R5–R11, grupa `pal_broj`) |
| **`PAL-ODL-12`** | **C002 ne prima izdanje sa 2b** (operater, 04.11.2026). Migrirana sveska C002 (izveštaji i ručne uplate kooperantima, bez pisanja paleta) ostaje na izdanju pre 2b; 2b se meri na fixture-u. Izlazna kapija C002 iz plana §14.58 („Drugi review“) zato otpada, a ponovni prevod iz stare sveske nije potreban |
| **`PAL-ODL-13`** | **Rez 2b ide kao dva PR-a** (operater, 04.11.2026: „Dva PR-a“). **2b-1:** kapija pisca `PAL-INV-01` (prazan, nepostojeći i tuđ `PrijemnicaStavkaID` i pogrešna klasa se odbijaju), C1 prijavljuje aktivnu paletnu stavku bez ključa, fixture i seed-ovi nose ključ linije; kolone ostaju. **2b-2:** brisanje `BrojPrijemnice`, `BrojZbirne` i `ZbirnaRoditeljID` sa paletne stavke, njihovih pisaca i konstanti. Prvo se podaci dovode u stanje u kome svaki red nosi ključ, pa se tek onda seče trag (plan §14.60) |

Podela S8 na rezove je prihvaćena po nezavisnom review-u (plan §14.58, „Review i odluke").

### Predlozi kako su postavljeni (pre odluka)

| ID | Pitanje | Preporuka | Šta zavisi od odgovora |
|---|---|---|---|
| **`PAL-ODL-01`** | Premisa: da li paletizaciju i preradu danas stvarno koristite, kod kog klijenta i koliko često? | — (pitanje, ne predlog) | koliko rigoroza ide u S8: ručna provera štampe i PDF-a, i da li `PAL-ODL-07` uopšte ima korisnika |
| **`PAL-ODL-02`** | Količina: da li je paletizacija **potpuna** (`PAL-INV-02`, jednakost), ili deo stavke sme da ostane van paleta? | **jednakost**, jer je to danas konstrukcija pisca; A3 prelazi na ključ stavke | pravilo pisca i A3 u rezu 2; `PAL-ODL-06` |
| **`PAL-ODL-03`** | Kolone paletne stavke: `BrojPrijemnice`, `BrojZbirne`, `ZbirnaRoditeljID` | **brišu se sve tri**; `PrijemnicaID` ostaje labela dokumenta (kao na liniji fakture, §4.5); `PrijemnicaStavkaID` je jedina veza i obavezan je. `ZbirnaRoditeljID` se izvodi kroz prijemnicu, a dok postoji, drži ga skriven drugi pisac (`ReassignPrijemnicaToZbirna_TX`) | obim reza 2 |
| **`PAL-ODL-04`** | Alt+F8 makroi koji menjaju palete i preradu po broju i tekućoj godini (`ClosePaleta_Prompt`, `StornoPaleta_Prompt`, `SavePrerada_Prompt`, `StornoPrerada_Prompt`, `PaletaAdjust_Prompt`) | četiri koja ekran Palete pokriva po ID-u se **brišu**. `PaletaAdjust_Prompt` ostaje, jer je on put „uradi kasnije": broj se razrešava u **tačno jedan** aktivan dokument, a dvosmislen broj se odbija i imenuje oba kupca | rez 1 |
| **`PAL-ODL-05`** | Siročad klase koju nova prijemnica nema: poruke i `MAPA` B-030 obećavaju „Skini stavke", a ta radnja ne postoji od S6 koraka 7 | **A:** radnja „Skini palete" u Oporavku nad izabranim redom (po `PrijemnicaID` stornirane prijemnice), kroz isto telo kao storno (`SkiniPaletneStavke`), sopstvena transakcija; prerađena paleta se odbija. **B:** samo ispraviti tekst; izlaz ostaje „Storniraj paletu", koji skida i su-stanare | rez 1 (A) ili tekst (B) |
| **`PAL-ODL-06`** | Storno palete koja nosi robu **aktivne** prijemnice (danas se skida sve, i su-stanari) | **odbiti** po imenu dok paleta nosi robu aktivne prijemnice; put je storno ili ispravka prijemnice, odnosno Oporavak. Dozvoljeno ostaje nad paletom bez aktivne robe | rez 2; doslednost sa `PAL-ODL-02` |
| **`PAL-ODL-07`** | Sledljivost na paleti: C-040 (ekran SLEDLJIVOST, obrisan u S1b-3b) i spisak kooperanata na paletnom i preradnom listu (prazan od S1b-1, C-050/C-051). Paleta nosi deo stavke, a stavka je zbir više otkupa, pa je paleta → kooperanti N:M bez alokacije: tačno se može reći samo „kooperanti od kojih je stavka sastavljena" (nadskup) | **prebaciti u S9**: isti graf sledljivosti, S8 daje tačan `PrijemnicaStavkaID`. Paletni i preradni list do S9 izlaze sa zaglavljem i bez spiska, kao danas | da li S8 ima rez 3 |
