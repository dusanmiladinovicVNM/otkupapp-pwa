# Refaktor: imenovani upis umesto pozicionog

> **Ulaz za svaku sesiju koja radi na ovom refaktoru.** Cilj: nijedan pisac ne zna poziciju kolone.
>
> **Verzija 4.** Verzije 1–3 su imale greške koje bi, da su izvršene, **ubile aplikaciju** (§1.1).
> Ovaj dokument je prošao tri nezavisne kritike i presudu nad njima (29.10.2026). Merenje je na
> `b552892` (PR #446). Ponašanje u Excelu je **NEVERIFIKOVANO** — `run_vba` traži Windows + Excel.

---

## 1. Šta je prethodna verzija pogrešno tvrdila

### 1.1 Kapija napisana po staroj §5 ubija svaku živu svesku

Stara §5 je u negativne slučajeve stavila *„sveska uža **ili šira** od kanona"* kao nešto što mora
dići grešku. **To je obrnuto od istine.**

- `modSchema.PrefiksNeslaganje:557` pada **samo** kad je `lo.ListColumns.count < kolone.count`.
- `modSchemaGuard.TabelaBrojKolona:303` vraća **živu** širinu (`lo.ListColumns.count`), ne kanonsku.
- **Širu svesku pravi sam projekat, na svakom startu:** `modSetup.EnsureRuntimeSchema:1221-1236` zove
  `EnsureColumnOnTable` nad `TBL_KULTURE`, `TBL_FAKTURA_STAVKE`, `TBL_VRSTA_GP`, `TBL_PRERADA`, a
  `:2250` dodaje **na kraj**. `.claude/rules/testovi.md` §9 to izričito zove legitimnim: *„one idu NA
  KRAJ. Legitimne su — samo ih kanon još ne drži."*

Kapija koja meri naspram kanona bi fail-closed obarala **svaki upis na svakoj živoj svesci**, a
`vba_check` i CI bi ostali zeleni — jer se to vidi samo u Excelu.

**Ispravno pravilo:** arnost se meri naspram **`lo.ListColumns.count` žive tabele**. Šira živa sveska
**mora da prođe**, a višak kolona dobija `Empty`. Pada **samo** sveska kojoj **kanonska** kolona
**fali** — i to već radi `RequireColumnIndex`.

### 1.2 Merenje je bilo pogrešno u sva tri prethodna oblika

Verzija 1 je tvrdila 163 mesta, v2 36/16/14, v3 39/18/21. **Sve troje netačno.** Uzrok je isti:
klasifikacija grep-om nad oknom linija ne vidi helper definisan daleko, ne razlikuje `Array(...)`
**imena kolona** od `Array(...)` **vrednosti**, i ne vidi fail-open granu.

**Ovo se ne sme meriti grep-om.** Potreban je parser koji spaja prelomljene redove, razrešava
helper i gleda da li je prvi `Array` imena ili vrednosti. Taj alat je **deo posla** (§4, P0).

### 1.3 Klasifikacija je trojna, ne binarna

Postoji treća grupa koju prethodne verzije nisu videle: **imenovan upis koji nepostojeću kolonu tiho
preskače.**

```vb
' modPaletniList.bas:1352-1359 -- PalAppendRow
n = lo.ListColumns.count
ReDim rowData(0 To n - 1)                               ' puna ZIVA sirina
idx = GetColumnIndex(tblName, CStr(cols(i)))            ' po IMENU
If idx >= 1 And idx <= n Then rowData(idx - 1) = vals(i)  ' <- FAIL-OPEN
```

Prvi `Array(...)` na pozivnom mestu su `COL_PALS_*` — **imena kolona**, ne vrednosti. Pozivna mesta
na preskakanje **računaju**: `:1321` *„PalAppendRow nepoznatu kolonu TIHO preskače, pa kolonu traži
RequirePaletaStavkaSchema pre ijednog upisa"*, `:2874` *„inače PalAppendRow tiho preskoči upis
bruto/paleta/ambalaza → 0 u sažetku paletnog lista"*.

### 1.4 `StampRowAudit` je vlasnik audit repa, i prethodna verzija ga nije pominjala

`modDataAccess.bas:294` zove `StampRowAudit` **posle** pozicione petlje. `:393-402`: `Created*` se
piše **samo ako je ćelija prazna** („ne gazi eksplicitnu vrednost"); `Modified*` uvek.

Pisci koji danas šalju kraći niz **ne diraju** audit rep, pa ga sistem ispuni. „Zaokruživanje reda na
punu širinu" bi upisalo vrednost pisca i **tiho ugasilo sistemski trag**. `modCenovnik.bas:132` već
danas piše `Now` u `CreatedAt`, pa ta grana tamo već okida.

**Pravilo: imenovani graditelj NE adresira `CreatedAt`/`CreatedBy`/`ModifiedAt`/`ModifiedBy`.**
Vlasnik je `StampRowAudit`; te kolone ostaju `Empty`. Nosi tvrdnju u `modTest`.

### 1.5 „Okidač ne dolazi skoro" je netačno

Stara §8 je prevod pisaca rangirala nisko jer *„nijedna od tih tabela nije u planu skraćivanja"*.
`docs/STANJE_REFAKTORA.md:39` kaže suprotno: **`S7 faktura · S8 palete · S9 sledljivost`** čekaju.

Istorija kanona pokazuje obrazac: **brisanje iz sredine se dogodi tačno kad se vrti korak te tabele.**
`7d2fb98` (27.09, S3-ostatak, `tblOtpremnica`) · `260ffd8` (01.10, `tblAmbalazaDokument`) · `d890814`
(13.10, `tblAmbalaza`) · **`cc3e294` (27.10, S6 korak 8 — obrisani `Kolicina`/`Cena`/`KolAmbalaze` iz
SREDINE `tblPrijemnica`)**.

**Deset od 14 pozicionih mesta živi u `tblFakture`/`tblFakturaStavke`/`tblPaleta`/`tblPaletaStavka`/
`tblPrerada`/`tblPreradaStavka`/`tblUtovar`/`tblUtovarStavke`** — tabelama S7 i S8.

---

## 2. Merenje (prošlo adversarnu proveru)

**41 produkcioni poziv** `\w*AppendRow` sa `TBL_` konstantom (+ 20 u test modulima):

| Grupa | Mesta |
|---|---|
| **Imenovano fail-closed** — puna živa širina, `RequireColumnIndex` | **21** |
| **Imenovano FAIL-OPEN** — puna širina, nepoznatu kolonu tiho preskače | **6** |
| **Pozicionо** — goli `Array(...)` vrednosti | **14** |

**Pozicionih 14, po fajlu:** `modUtovar:1270,1285,1593,1603` · `modNovac:230,351` ·
`modSEFPersistance:355,501` · `modFaktura:328,345` · `modAgrohemija:161` · `modCenovnik:137` ·
`modStornoContext:79` · `modStornoZurnal:95`

**Fail-open 6:** `modPaletniList:1281,1323,2934,2951` · `modAgrohemija:397` · `modKooperant:140`

`modDokumenta` (11) i `modOtkup` (2) su **imenovani fail-closed** — jezgro je završeno.

### 2.1 Obrazac se ne projektuje, postoji

`modOtkup.BuildOtkupHeaderRowData` radi tačno to, i **dokazan je stvarnim događajem**: `018be597`
(S1d, 17.09) obrisao je 8 kolona iz **sredine** `tblOtkup` (37→29), a `modOtkup.bas` **nije dirnut
nijednom linijom**. `SetRowValueByColumn` (`modSchemaGuard:281`) se koristi 89 puta.

`.claude/rules/podaci-i-config.md` to već propisuje: *„Za polja čiji redosled nije siguran koristi
upis po imenu"*, uz napomenu da *„kapija štiti od zatečene sveske, ne od novog koda koji pogrešno
složi niz"*.

### 2.2 Baza niza nije uniformna

`SetRowValueByColumn:289` piše `rowData(colIndex - 1)` → **0**-bazno. `modUtovar:157`
`ReDim rowData(1 To lo.ListColumns.count)` → **1**-bazno. Kapija **mora normalizovati `LBound`**.

### 2.3 Dva SEF pisca imaju ZAKUCANU širinu

`modSEFPersistance.bas:331` `Dim rowData(1 To 20)`, `:488` `Dim rowData(1 To 9)` — literali, ne
`TabelaBrojKolona`. Danas su tačno jednaki kanonu, pa ih **kapija arnosti nikad ne okine**. Kad se
tabela proširi, nova kolona ide na kraj, literal ostaje, poslednja kolona nikad ne dobije vrednost, a
kapija propusti. `tblSEFEventLog` uz to **nije** gejtovan `SchemaReadyOrFail`-om.

### 2.4 Tri pozicionа pisca nisu gejtovana ni `SchemaReadyOrFail`-om

`modCenovnik:137` (modul nije ni u spisku pozivaoca) · `modNovac:351` za `tblPartnerMap` (kapija na
`:71` prosleđuje samo `TBL_NOVAC`) · `modSEFPersistance:501` za `tblSEFEventLog`. Za ta tri
reorderovana sveska **danas** tiho piše u pogrešne kolone.

---

## 3. Pre-flight verdikt (ispravljen)

| Osa | Status | Dokaz |
|---|---|---|
| `DOMAIN` | **GAP** — samo `modFaktura` i `tblNovac` | `UPIS-KONV-01` je **mrtva**: `tblAmbalaza` u kanonu **nema** kolonu `Stornirano`, a `tblAmbalazaDokument.Stornirano` se piše eksplicitno (`modAmbalaza.bas:1461`). Ostaje `modFaktura` (`IzdatoStatus` prazan) i — što stara verzija nije videla — `tblNovac` |
| `IDENTITY` | `N/A` | Nijedan identitet se ne menja |
| `CARDINALITY` | `N/A` | Jedan niz = jedan red |
| `INVARIANTS/OWNER` | **GAP** | Ne „PROVEN": `KNOWN_ISSUES.md:55` **KI-029 OPEN** — `who_writes.py:291` čita `WRITE_OWNERSHIP.json` golim `json.load`, duplikat ključa uzima **poslednji**. Registar na kom osa stoji može tiho izgubiti vlasnika |
| `WRITERS` | **PROVEN** | §2. Ime novog mutatora mora da **završava** na `AppendRow`/`UpdateCell`/`DeleteRow` (`who_writes.py:81` slobodan prefiks, strog sufiks; token je `UpdateCell`, **ne** `UpdateRow`), i tabela ostaje **prvi** argument kao `TBL_` konstanta |
| `DOWNSTREAM` | **PROVEN** | Nijedno mesto ne čita red literalnim indeksom — svi po imenu, i zato je kvar nevidljiv: `modSEFMapper` verno pošalje pomerenu vrednost **u SEF**. Žurnal je drugi potrošač i **ima čitaoca na startu** (`modMain.bas:217 CheckJournalForRecovery`) |
| `CAPABILITY` | **GAP** za fail-open grupu | Prevod fail-open → fail-closed menja „paleta sa nepotpunim sažetkom" u „paleta se ne može napraviti". Nosi `MAPA_SPOSOBNOSTI.md` C-046/C-048/C-049, a `STANJE_REFAKTORA.md:14` kaže da se mapa sposobnosti **apsolutno čuva** |
| `ACCEPTANCE CONTRACT` | **GAP** | §5 |
| `PLATFORM` | **UNMEASURED, i BLOKIRA** | Pun imenovani niz menja „nedirnuto" u `Empty` na svakoj koloni žive sveske; danas `Application.Min` klamp znači da rep **nikad** ne dobije upis. Šta `ListRows.Add:283` ostavlja u nepopunjenoj koloni nije izmerivo bez Excela. **Odluka pre P2** |
| `LANDING` | **GAP** | Stara verzija je merila nad **spojenim** PR-om. Otvoren je **#447** („Bilans ambalaže", nije draft) i dira `src-vba/modTest.bas`, `tools/sabotaza.py`, `docs/STANJE_REFAKTORA.md` — **tačno fajlove koje §5 obavezno dira** |

`EVENTS: N/A — refaktor sloja upisa.`

---

## 4. Rez

Jedno imenovanje kroz ceo dokument: **P0…P7**. (Stare oznake B/C/D ne postoje više.)

| PR | Sadržaj | Blokira ga |
|---|---|---|
| **P0** | **Alat koji meri** — parser koji spaja prelomljene redove, razrešava helper i razlikuje `Array` imena od `Array` vrednosti; daje trojnu klasifikaciju. Nijedna izmena VBA | — |
| **P1** | `modSetup.EnsureDataTable`: **prisutan kvar van plana.** Tri `Array`-a odstupaju od kanona **u sredini** (`tblPaleta` 20/23 od pozicije 6, `tblPaletaStavka` 13/18 od 3, `tblPrerada` 17/22 od 5). `:2206` postojećoj tabeli samo dopunjava, `:2222` nepostojeću pravi **tačno** u redosledu `Array`-a, `:2250` dodaje na kraj — redosled se **nikad** ne popravi. Dostupno sa Admin dugmeta (`modAdmin.bas:294`), i od tog momenta `modPaletniList.bas:165 SchemaReadyOrFail` **trajno** odbija paletizaciju. Kapija: Python poredi svaki `EnsureDataTable` `Array` sa kanonom | — |
| **P2** | **Asimetrija strogih čitalaca članstva zbirne.** Prenos šest tvrdnji iz `modDokumenta:5723-5758` (`AktivnoOtpClanstvoPoKanonu`: prazan roditelj 1331, prazno dete 1332, roditelj ne postoji tačno jednom 1334, dete 1335, dupli par 1333, dva aktivna 1305) u `modDokumenta:1598-1612` (`AktivnoClanstvoPoKanonu`, zbirna, 7 pozivaoca, ima **samo** 1260, a `:1601` red sa praznim ID-em **tiho preskače**). ~20 linija, obrazac postoji sprat niže | — |
| **P3** | **Integritet nad S6 vezama.** `tblPrijemnica.ZbirnaRoditeljID`, `tblPaletaStavka.PrijemnicaStavkaID`, `tblFakturaStavke.PrijemnicaStavkaID` — **nula provera**, a `Chk_B5:264` meri **labelu** `BrojZbirne`, pa prijemnica sa visećim `ZbirnaRoditeljID` prolazi ako je labela popunjena (`prj_brojzbirne 33 = prag 33`) | P2 |
| **P4** | **KI-036 + KI-029** — merila kojima se sve ostalo dokazuje. Dok `dokaz.py` može upisati zeleni marker nad nečistim izvorom, §5 ništa ne tvrdi | — |
| **P5** | `modJournaling`: **širina FAJLA, ne reda.** `:94 If Dir(filePath) = ""` — zaglavlje se piše jednom po tabeli po danu, `:98` po **živoj** širini. Dopuna reda „do žive širine" posle `EnsureRuntimeSchema` daje red **širi od zaglavlja u fajlu** → crash-recovery uvoz tiho pomeri svako polje. Danas je red kraći, pa je manjak **vidljiv**. Rešenje mora da reši fajl, inače je **gore od zatečenog** | P4 |
| **P6** | Prevod **14 pozicionih** pisaca — **preduslov za S7/S8** (§1.5). `tblFakture` piše **dva** modula poziciono (`modFaktura:328`, `modUtovar:1593`), pa prevod jednog ne rešava ništa. Uz to: SEF literali (§2.3) i tri negejtovana (§2.4) | P0, odluke §6 |
| **P7** | **Fail-open grupa (6)** — zaseban korak uz odluku, **ne mehanički**. Ili se svuda dodaje `Require*Schema` pre transakcije (`modPaletniList:1322` to već ima za `PALETA_STAVKA`, `:1281` za `PALETA` **nema**), ili se degradacija svesno ukida i upisuje u `MAPA_SPOSOBNOSTI` uz C-046/C-048/C-049 | `CAPABILITY` odluka |

**Kapija arnosti u `AppendRow` nije u rezu.** Stara P4 je bila nedefinisana na tri točke i uvodila
`newRow.Delete` — oblik koji `.claude/rules/podaci-i-config.md` zabranjuje („brisanje reda je razred
upisa"; golo `newRow.Delete` je nevidljivo A11 kapiji, mora kroz `RequireDeleteRow`), a `newRow` može
biti `Nothing` pa `Delete` u handleru baca 91 **iz** handlera. Uz to: 41 produkcioni + 20 testnih
pozivaoca je napisano na `<= 0`/`= 0`, a plan nije rekao šta `AppendRow` vraća posle kapije. Vraća se
kad P6 i P7 završe, kao zaseban pre-flight.

**Merilo napretka ne može biti grupa u `popis_citalaca.py`.** `:79-123` `GRUPE` su regexi nad
**sidrima** (`COL_*` konstante, imena procedura), a `Array(newID, brojDok, datum, ...)` ne nosi ni
jedno sidro. Uz to `:1048-1057` traži **tačnu jednakost** praga (pada i iznad i ispod), pa bi prag 21
nad 14 mesta dao exit 1 na prvom commit-u. Merilo je alat iz P0.

---

## 5. ACCEPTANCE CONTRACT

Stara §5 je citirala pet mrtvih ili zastarelih stvari: suite `imenovani-upis` **ne postoji** u
`tools/run_vba.py` (18 suita; `.claude/rules/testovi.md` §3: *„katalog SUITES je jedini izvor istine"*
i *„Nova suite mora biti gate"*); `RequireRedArnost` — **nula pogodaka** u repou (sabotirala je ime
koje nikad nije definisano); prefiks za `dokaz.py` ne postoji; „598 sabotaža" je danas **1040**;
„`modTest` 199 testova" je **288**; „34 tabele" je **36**.

**Zato je `ACCEPTANCE CONTRACT` = GAP i piše se po PR-u, ne za ceo refaktor.** Svaki PR nosi:
šta važi kad se završi · šta ostaje netaknuto · koji negativan slučaj pada **po imenu** · kojim
testom i u kojoj **postojećoj** suite (ili koja se dodaje u `SUITES` kao gate).

Baseline-ovi na dan 29.10.2026, izmereni: `vba_check` exit 0 „čisto (191 fajlova)", 1040 sabotaža ·
`gen_schema_module --check` exit 0, otisak `87AB4F9B` · `who_writes --check-ownership` exit 0,
36 tabela · `popis_citalaca --check` exit 0, 13 grupa · `modTest` očekivano 288/0.

**Negativni slučajevi koji moraju pasti** (ispravljeno po §1.1): red **uži** od žive širine · red
**širi** od žive širine · nepoznato ime kolone · red koji nije niz · **kanonska kolona koja fali**.
**Mora da PROĐE:** živa sveska **šira** od kanona (višak dobija `Empty`).

**Dokaz u oba smera** je obavezan (menja se checker), ali tek posle P4 — jer `KI-036` kaže da
`dokaz.py` može upisati zeleno nad sabotiranim izvorom.

---

## 6. Odluke koje čekaju operatera

| ID | Pitanje | Blokira |
|---|---|---|
| ~~`UPIS-KONV-01`~~ | **MRTVA** — `tblAmbalaza` u kanonu nema `Stornirano`; `tblAmbalazaDokument` ga piše eksplicitno (`modAmbalaza:1461`) | — |
| `UPIS-KONV-02` | `tblFakture.IzdatoStatus` ostaje prazan (`modFaktura:304-327`); konvencija „prazno = IZDATO". Eksplicitno ili izostaviti? | P6 |
| `UPIS-KONV-04` | **Novo** — `tblNovac` isto pitanje, stara verzija ga nije videla | P6 |
| `UPIS-KONV-03` | `modUtovar:1265` računa na tiho dopunjavanje. Šta ide u te kolone? | P6 |
| `UPIS-CAP-01` | Fail-open grupa: `Require*Schema` pre transakcije (čuva sposobnost) ili svesno ukidanje degradacije uz upis u `MAPA_SPOSOBNOSTI`? | P7 |
| `UPIS-PLATFORM-01` | Pun niz upisuje `Empty` u kolone koje danas ostaju nedirnute. Šta `ListRows.Add` tamo ostavlja? Traži sondu na Windows + Excelu | **P6** (ne „ne blokira") |

---

## 7. Preuzimanje grane

```bash
cd ~/Documents/GitHub/otkupapp-pwa
git fetch origin claude/great-hawking-91ti26
git checkout claude/great-hawking-91ti26
git pull origin claude/great-hawking-91ti26
python tools/vba_check.py
```
