# Ocena plana „idealni AgriX VBA codebase" + korigovani plan

- **Status:** Analiza / predlog smera.
- **Verzija:** v5 (2026-09-07). v5 ne menja nijedan zakljucak — **vadi brojke iz
  teksta** i ispravlja dve tvrdnje koje su vodile odluku, a bile netacne.
- **Brojke:** `docs/Architecture/ARH_SNIMAK.md` — generisan
  (`python3 tools/arh_snimak.py --out docs/Architecture/ARH_SNIMAK.md`).
  **Ovaj dokument namerno ne ponavlja brojeve.**
- **Odluke:** `docs/adr/0003-repository-granica-i-izuzeci.md`
- **Istorija verzija i sopstvenih gresaka:** §9

> **Zasto je podela ovakva.** v4 je bio tacan na dan pisanja i **netacan za jedan
> dan** — repo je u 166 commita obrisao legacy forme (posao koji je v4 zvao
> najvecim preostalim) i dodao dva nova pisca poslovnim tabelama (metrika koju je
> v4 zvao skoro resenom). Dokument je time gresio bas na dve tacke koje vode
> odluku, i to u suprotnim smerovima: slao bi coveka na zavrsen posao, a sklanjao
> ga sa onog koji gori.
>
> Zakljucci su izdrzali tri premeravanja. Brojke nisu izdrzale jedan dan. Zato
> **zakljucci ostaju ovde i pisu se rukom, a brojke se generisu** i citiraju iz
> snimka.

---

---

## 0. Kratak sud

Smer originalnog plana (Host → Presentation → Application → Domain → Repository →
Infrastructure, Sync sa strane) je **tacan i vredi ga zadrzati**.

Dijagnoza trenutnog stanja u njemu **nije bila tacna**, i redosled koji je iz nje
sledio bio je pogresan. Plan je pretpostavljao da Application sloj ne postoji i da
UI pise u tabele; nijedno nije bilo tacno. Istovremeno nije video dve stvari koje
stvarno kostaju.

| | Ocena |
|---|---|
| Ciljna slika | **dobra** — zadrzati |
| Dijagnoza zatecenog stanja | **slaba** — devet materijalnih gresaka (§2) |
| Redosled implementacije | **los** — invertovan (§4) |
| Fizicka organizacija (`00_Host/`…) | **odbaciti** — skupo, bez efekta |
| CI dependency rules | **najbolja stavka** — sprovedena prva (§4, PR1) |

---

## 0a. Kako je ovaj dokument gresio — i sta iz toga sledi

Cetiri puta je merenje oborilo zakljucak koji je zvucao tacno. Tri puta moj, jednom
autorov. To nije uvod u skromnost nego **razlog za podelu na plan i snimak**:

1. **v1:** „Application sloj nedostaje." Ne nedostaje — `modOtkupUnos`, `modDokUnos`,
   `modNovacUnos`, `modAgroUnos` nose oblik *command → validate → execute*.
2. **v2:** „Svih `*_TX` je Application sloj." Presiroko — vecina je puki
   transakcioni omotac oko blizanca bez `_TX`, dakle mehanizam a ne namera.
3. **v3:** „Toliko-i-toliko fizickih pisaca nad `tblOtkup`." Pogresan signal —
   brojka je dolazila iz `WHO_WRITES.md`, koji sabira `tx` (snapshot) i `direct`
   (stvarni upis). Za write-gateway metriku vazi samo drugi.
4. **v4:** dve tvrdnje koje vode odluku, obe pogresne **za jedan dan**:
   - legacy duplikacija oznacena kao „najveci preostali posao, jedina stavka koju
     nijedno premeravanje nije smanjilo" — **zavrsena je** dok je dokument stajao;
   - vlasnistvo upisa oznaceno kao „skoro reseno" — **pogorsava se**, novi moduli
     su se upisali u tabele koje su vec imale svog pisca.

Cetvrta greska je najskuplja jer je dvosmerna: dokument bi poslao coveka na
zavrsen posao, a sklonio ga sa onog koji gori.

**Zato od v5 vazi:** zakljucci se pisu rukom u ovom fajlu, brojke se generisu u
`ARH_SNIMAK.md`. Tvrdnja tipa „X ih ima toliko" ovde vise ne stoji — stoji samo
„X je ovakve prirode, i evo zasto".

---

## 1. Sta je mereno i kako

Sve mere su staticke, nad `src-vba/`, bez Excela — dakle rade i u Linux sesiji.
Vrednosti: `ARH_SNIMAK.md`. Ovde je samo **sta svaka mera znaci**, jer se to ne
menja kad se brojka promeni.

| Mera | Sta tvrdi | Zasto bas ona |
|---|---|---|
| **Fizicki pisci po tabeli** | koliko modula zove `AppendRow`/`UpdateCell` nad tabelom | metrika Repository faze. **Nije** isto sto i „ko tabelu poslovno menja" — poslovnih pisaca sme biti vise, fizickih jedan |
| **`TBL_`/`COL_` u sloju prikaza** | koliko ekran zna o strukturi baze | preostali dug Presentation sloja; upis je vec nula |
| **`upis` u sloju prikaza** | zove li ekran `AppendRow`/`UpdateCell`/`GetNextID` | mora ostati 0; cuva `SLOJ_UPIS` u `vba_check` |
| **`*_TX` omotaci** | koliko `_TX` procedura samo otvara transakciju oko blizanca | razlikuje *mehanizam* od *use-case-a* |
| **`Public` blizanci** | koliko poslovnih upisa je pozivljivo mimo transakcije | vrata pored granice; danas ih niko ne koristi pogresno, ali to nije masinski provereno |
| **`BeginTx` vs `AddTableSnapshot`** | drzi li invarijantu „ko otvara transakciju, taj deklarise tabele" | osnov ADR-0003 tacke B. Treci broj mora biti **0** |

Merenje `TBL_`/`COL_` broji **pojave, ne linije** — jedna linija zna da nosi
tabelu i tri kolone. Raniji rucni `grep -c` u v1–v4 bio je linijski i zato uvek
manji; posao nizvodno nije „obidji linije" nego „zameni svaku referencu".

---

## 2. Materijalne greske u originalnom planu

### M1 — „Application sloj najvise nedostaje"

Netacno. Postoje `modOtkupUnos`, `modDokUnos`, `modNovacUnos`, `modAgroUnos` u
obliku `Novi*Unos` → `*Validiraj` → `*Upisi` — tacno *command → validate →
execute* koji plan predlaze kao `clsKreirajOtkupCmd` → `modAppOtkup.Kreiraj`.

Precizna formulacija: sloj **postoji ali nije koherentan ni ogranicen**. Dva
imenovanja (`*Unos` i `*_TX`), od kojih drugo u vecini slucajeva oznacava samo
transakciju.

**Ispravka:** ne uvoditi `modApp*` pored toga — to je trece ime za isti sloj.
Preimenovati postojece samo uz izmenu koja ionako dira fajl. Sufiks `_TX`
zadrzati: nosi informaciju koju `modApp` prefiks ne nosi, i na njega se vezuje
provera `TX_VRATA`.

### M2 — Repository na petom mestu

Repository je jedini korak koji direktno obara metriku vlasnistva upisa. Sve pre
njega je preraspodela imena.

**Precizacija cilja:** nije „jedan poslovni pisac po tabeli". Poslovnih pisaca i
dalje treba vise — `modStorno` stornira, `modMasterSync` uvozi, `modDokumenta`
vezuje. Cilj je **jedan fizicki write gateway**:

```
modOtkup      -> RepoOtkup.Insert
modStorno     -> RepoOtkup.MarkStornirano
modMasterSync -> RepoOtkup.InsertImported
modDokumenta  -> RepoOtkup.SetOtpremnica
                        |
                        v
                  modRepoOtkup     <- JEDINI koji sme AppendRow/UpdateCell
                        |
                        v
                  modDataAccess
```

### M3 — „Domain tests — no workbook"

VBA nema host-free runtime; `run_vba.py` trazi Windows + Excel + `pywin32`.
Domain test bez workbook-a je **fizicki neizvodljiv**.

**Ispravka:** cilj je test **bez table fixture-a i bez `clsTransaction`**. To je
razlika izmedju testa od desetina milisekundi i testa od nekoliko sekundi, i to
je stvaran dobitak.

### M4 — Repository kao klase

Plan u §5 odbacuje interface-e, a u §20 trazi **fake repository**. U VBA bez
`Implements` fake se ne moze ubaciti — `clsOtkupRepository` je konkretan tip.

**Ispravka:** `modRepo*` **standardni moduli**.

Obrazlozenje iz v1 ove ocene bilo je pola netacno i ispravlja se: standardni modul
**ne daje** isti fake seam kao klasa — nije objekat, ne moze se injektovati. Tacan
razlog je drugi: **fake seam trenutno ne vredi dodatnu kompleksnost.** Standardni
modul daje ono zbog cega Repository i uvodimo — ekskluzivno vlasnistvo nad
`TBL_`/`COL_` i upisom, masinski proverivo — po ceni nula lifetime managementa.
Put ka `Implements` ostaje otvoren za 3–4 aggregate-a, uz uslov **demonstrirane
potrebe**, ne estetike.

### M5 — Fizicka taksonomija `00_Host/` … `90_Tests/`

VBA namespace je **ravan**, a `.frm` mora u commit sa `.frx` parom
(`CLAUDE.md` §3). Masovno preimenovanje modula: veliki diff, rizik od kolizije
`Public` imena, nula uticaja na dependency graf.

**Ispravka:** zadrzati flat `src-vba/`. Sloj se izrazava **prefiksom**
(`modScr*`, `modApp*`, `modDom*`, `modRepo*`, `modQry*`, `modSync*`), a granica se
namece u `vba_check`. Prefiks je vec ustaljen i radi.

### M6 — `modOtkupUI` kao gotova ljuska

Ljuska je `modUiScreens` — mali, cist registry bez ijedne `TBL_` reference.
`modOtkupUI` je ljuska **plus** najveca pojedinacna kolicina ekranske data-logike
u codebase-u (KPI agregacije, `FillComboDisplayID`, partner mape, liste).

**Ispravka:** `modOtkupUI` je klijent Query sloja br. 1, ne „gotovo".

### M7 — Plan ne slece u canonical dokument

`ARCHITECTURE_REFERENCE.md` §0.3: *„Ako nesto nije navedeno u ovom dokumentu, ne
smatra se canonical arhitekturom."* Plan koji ne udje tamo nije obavezujuci.

**Ispravka:** svaka prihvacena faza zavrsava upisom u `ARCHITECTURE_REFERENCE.md`
+ red u `ARCHITECTURE_CHANGELOG.md`. Odluke o smeru idu kao ADR — ADR-0003 je
prvi takav.

### M8 — Nema izlazne metrike ni po jednoj fazi

Deset koraka, nijedan merljiv kriterijum „gotovo".

**Ispravka:** svaka faza u §4 nosi metriku iz `ARH_SNIMAK.md`, pa se „gotovo"
proverava komandom a ne procenom.

### M9 — Repository moze da se izrodi u „lepsi DataAccess"

*(Nalaz autora plana, prihvacen — najozbiljnija zamerka korigovanom planu.)*

```vba
' LOSE -- ovo je UpdateCell sa prefiksom
RepoOtkup.UpdateColumn otkupID, COL_OTK_OTPREMNICA_ID, otpID
RepoOtkup.SetField otkupID, "BrojZbirne", broj
```

```vba
' DOBRO -- namera, ne mehanika
RepoOtkup.LinkToOtpremnica otkupID, otpID
RepoOtkup.AssignZbirna otkupID, brojZbirne
RepoOtkup.MarkStornirano otkupID, razlog
```

Razlika nije stilska: kod generickog API-ja invarijanta „kad se postavlja
`BrojZbirne`, proveri ownera" **nema gde da zivi** i ostaje razbacana po
pozivaocima — tacno stanje koje Repository treba da ukine.

Sprovedeno kao provera `REPO_API` u `vba_check` (§4, PR1).

---

## 3. Sta u planu ostaje netaknuto

1. **Ciljni dijagram** — tacan i za VBA izvodljiv.
2. **CQRS-lite:** read ne mora kroz Domain. Jedini deo plana koji pravilno
   oslovljava stvarni Presentation dug.
3. **Business ID ≠ broj dokumenta** — delom kodifikovano u ADR-0001/0002.
4. **`AllocateNewNumber` / `ObserveExistingNumber`** — `modBrojevi` danas radi pet
   strategija u jednom modulu; razdvajanje *alociraj* vs *upamti vidjeno* je cista
   dobit, narocito za naknadni papirni unos.
5. **CI dependency rules** — najbolja i najjeftinija stavka.
6. **„Ovo nije rewrite nego reorganizacija ownership-a"** — tacan okvir.

---

## 4. Plan

Princip: **prvo ono sto se namece masinski i obara merljivu metriku; imenovanje
poslednje.** Svaka stavka nosi metriku iz `ARH_SNIMAK.md`.

### PR1 — `SLOJ`, tvrdi deo  ✅ URADJENO

Cetiri provere u `tools/vba_check.py`, **nula nalaza** nad zatecenim izvorom,
**bez baseline fajla**:

```
SLOJ_UPIS     modScr* / modOtkupUI  ne smeju AppendRow / UpdateCell / GetNextID
SLOJ_UZVODNO  modDataAccess         ne sme modScr* / modApp* / modDom* / frm*
REPO_TX       modRepo*              ne sme BeginTx ni AddTableSnapshot  (ADR-0003 B)
REPO_API      modRepo*              ne sme genericki Public API          (M9)
```

Ne ciste nista — **zakljucavaju stanje koje je vec postignuto.** `REPO_*` su
prazne dok `modRepo*` ne postoji; to je i bio razlog da udju sada.

*Metrika:* kolona `upis` u snimku ostaje 0 u svakom redu.

### PR2 — kapija nad brojem pisaca  ← SLEDECE, i najhitnije

`who_writes.py --max-writers`, prag na **zatecenim vrednostima**, exit 2 iznad.

Ne cisti nista — sprecava dalje pogorsanje. Ovo je jedina stavka u planu koja
gasi **aktivno** pogorsanje: novi moduli se upisuju u tabele koje vec imaju svog
pisca, i to se ne vidi dok neko ne pogleda — a po `CLAUDE.md` §2 pogleda se tek
kad pravilo upisa vec treba menjati.

*Metrika:* „tabela sa vise od jednog pisca" u snimku prestaje da raste.

### PR0 — `SyncControl` write ownership

Dve putanje pisu isti tab, jedna whole-tab replace-om. `modStanicaLock` je vec
fail-closed, `modGoogleSyncOrchestrator` nije. Data-safety, ne arhitektura.

`modSyncControl` kao jedini vlasnik; `TryReadSyncControlAsDict` +
`ApplySyncControlUpdates` sele se iz `modStanicaLock`.

*Metrika:* „modula koji pominju `SyncControl`" → 1. *~1 dan.*

### F1 — Repository, samo gde treba

Vecina tabela vec ima tacno jednog fizickog pisca — tamo nema sta da se radi.
Rade se samo one iznad jedan, po snimku. Semanticki API (M9), TX-neutralan
(ADR-0003 B).

> **ADR-0003 A:** `modSetup` i `modMigracija` su **imenovan izuzetak** — bootstrap
> admin naloga i jednokratna migracija nisu poslovni dogadjaj. Prihvacena cena:
> tabele koje pise `modSetup` ostaju visepisacke.

### F2 — Ekranska disciplina, kao racva a ne projekat

Ne uvodi se sloj. **Primenjuje se obrazac koji dva ekrana vec koriste**
(`modScrSledljivost`, `modScrBankaNalozi` — nula `TBL_`/`COL_`, podatak traze od
poslovnog modula, a kolonsku specifikaciju sa sirinama drze kod sebe).

> **Granica koju ne treba preci:** Query vraca **podatak**, ne izgled.
> `DocumentListRow{Datum, Broj, Partner, Kolicina, Status}` — ne sirinu kolone,
> bold ili redosled. Inace se coupling samo premesti sa baze na grid, i `modQry*`
> postane drugi `modScr*`.

**Kao racva, ne kao projekat:** zamrznuti broj `TBL_`/`COL_` po fajlu i dozvoliti
mu **samo da opada**. Migracija se onda desava kao nusprodukt rada na ekranima.
Pri ovoj brzini isporuke refaktor-projekat gubi trku sa feature radom.

*Metrika:* zbir `TBL_`+`COL_` u sloju prikaza monotono opada.

### F3 — `TX_VRATA`

Blizanac bez `_TX` mora biti `Private` ili pozivan iz transakcije. Danas su ta
vrata otvorena ali ih **niko ne koristi pogresno** — sto nije masinski provereno,
pa je to provera koja nedostaje, ne bag koji treba gurati.

### F4 — Domain gde ima ROI, sync adapter, numbering

- `modDomStorno` iz `modStornoImpact`/`modStorno`; `modDomDokument` iz
  `modDokumentInvariant`; `modDomOtkup` (bruto→neto, klase, ambalaza, cena).
  `modNovac`, `modFaktura`, `modDokumenta` **ne dirati** bez povoda.
- `modMasterSync`: `parse DTO` odvojiti od upisa; import zove `modOtkupUnos`.
- `modBrojevi` → `AllocateNewNumber` / `ObserveExistingNumber` + `NumberRegistry`.

**Domain je najnizi prioritet u planu.** Pri ovoj brzini promena odlozio bih ga
bez roka.

### F5 — Imenovanje, oportunisticki

Samo uz izmenu koja ionako dira fajl. Nikad kao zaseban rename commit.

### Odbaceno

| Stavka | Razlog |
|---|---|
| `00_Host/` … `90_Tests/` folderi | VBA namespace je ravan (M5) |
| `clsOtkupRepository` i sl. | fake seam se ne dobija ni klasom bez `Implements`, a sam seam jos ne vredi kompleksnost (M4) |
| `clsKreirajOtkupCmd` command klase | `Object`/`Dictionary` payload koji `*Unos` vec koristi radi isto |
| „Domain tests — no workbook" | fizicki neizvodljivo u VBA (M3) |
| Novi `modApp*` pored `*Unos` | trece ime za isti sloj (M1) |
| **Legacy konvergencija** | **zavrseno na `main`-u** — `frmOtkup` i `frmDokumenta` uklonjeni, ostala jedna forma |

---

## 5. Sazetak: sta je AgriX-u zaista potrebno

> **AgriX-u ne treba jos slojeva. Treba mu ownership nad slojevima koji vec
> fakticki postoje.**

| Sloj | Postoji? | Sta stvarno nedostaje |
|---|---|---|
| Presentation write separation | **da** | nista — samo zakljucati pravilom (uradjeno, PR1) |
| Presentation read separation | **delimicno** | `TBL_`/`COL_` u ekranima; obrazac postoji, primeniti ga |
| Query obrazac | **da**, dva radna ekrana | prosiriti na ostale |
| Application-ish sloj | **da** | koherentnost; omotaci mesaju mehanizam i nameru |
| Transaction boundaries | **da** — invarijanta bez izuzetka, v. snimak | vrata pored granice, masinski neprovereno |
| **Repository** | **vecina tabela vec** | manjina iznad jedan — **i raste** |
| Sync idempotency | **da** | jedan ulaz umesto direktnog upisa |
| Enforcement granica | **delimicno** (PR1) | kapija nad brojem pisaca (PR2) |
| Legacy jedinstvenost | **da** | **zavrseno** — legacy forme uklonjene, ostala jedna |

Jedina stavka koja se **pogorsava** je vlasnistvo upisa. Sve ostalo ili stoji ili
se popravlja samo od sebe kroz redovan rad.

---

## 6. Sta ovo NE resava

- **Ne popravlja nijedan postojeci bag.** Nijedan nalaz u ovoj analizi nije bio
  demonstriran bag — bila su to nedostajuca merenja.
- Ne ubrzava nista, ne dodaje funkcionalnost.
- Ne smanjuje `modTest.bas` ni potrebu za Windows/Excel runnerom.
- Ne dira `.frx` — nove kontrole i dalje idu runtime-om.
- **Ne resava schema drift.** Sema tabela ostaje izvor istine po instalaciji
  (`CLAUDE.md` §3). Repository cak **povecava** vaznost te provere, jer
  centralizuje pretpostavke o kolonama na jedno mesto.

---

## 7. Preporuka

**PR2 (kapija nad brojem pisaca)** je najhitnija stavka — jedina koja gasi aktivno
pogorsanje, i posao od jednog dana.

**PR0 (`SyncControl`)** odmah posle: data-safety, nezavisno od svega ostalog.

**F2 kao racva, ne projekat.** Najbolji odnos dobitka i rizika: read-only, greska
se vidi na ekranu, obrazac je dvaput dokazan u repou.

**F1 je vredan ali uzi nego sto su v3/v4 tvrdili.** Vecina tabela je vec resena.

**Ono sto se ne sme raditi prvo:** Domain extraction. Najnizi ROI u planu.

### Sire od ovog plana

v4 je bio tacan na dan pisanja i netacan za jedan dan. To nije mana plana nego
**mana formata**: u repou koji isporucuje ovom brzinom, faza koja traje nedelju
dana ne stigne da se izvrsi pre nego sto joj se pomeri osnova.

Zato: **kapije umesto faza.** Kapija je jedan dan posla, ne zastareva, i pretvara
svaki buduci PR u napredak — dok faza trazi da se svet ne pomeri dok je radite.

---

## 8. Reprodukcija

```bash
# sve brojke koje ovaj dokument citira
python3 tools/arh_snimak.py --out docs/Architecture/ARH_SNIMAK.md

# vlasnistvo upisa, sa oba signala (tx + direct)
python3 tools/who_writes.py --out docs/DOMEN/WHO_WRITES.md

# granice slojeva
python3 tools/vba_check.py
python3 tools/vba_check.py --self-test
```

---

## 9. Istorija verzija

| Verzija | Sta je donela | Sta je od nje oboreno |
|---|---|---|
| **v1** | originalni plan i ciljna slika | dijagnoza zatecenog stanja (M1–M9), redosled, folder taksonomija |
| **v2** | prvo premeravanje: Application postoji, UI ne pise, legacy duplikacija | „svih `_TX` je Application"; „jedan pisac po tabeli"; obrazlozenje za `modRepo*` |
| **v3** | korekcije autora, `_TX` klasifikacija, M9, preplitanje faza | broj fizickih pisaca (pogresan signal), velicina Repository faze, „uvedi Query sloj" |
| **v4** | premereno posle 109 commita; `who_writes.py` popravljen | legacy kao „najveci preostali" (zavrsen za jedan dan), vlasnistvo upisa kao „skoro reseno" (pogorsava se) |
| **v5** | brojke izvucene u generisan `ARH_SNIMAK.md`; plan drzi samo zakljucke | — |

Metod je kroz svih pet isti i vredi ga zadrzati: **tvrdnja koja nije izmerena se
ne upisuje kao nalaz.** v5 dodaje drugu polovinu tog pravila — **izmerena tvrdnja
se ne prepisuje rukom**, jer merenje zastareva brze nego sto se dokument cita.
