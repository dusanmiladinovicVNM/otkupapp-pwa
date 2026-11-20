# Brzina: snimanje unosa, izmena dokumenata i prikaz u Izveštajima

Statička analiza VBA sloja, 18.11.2026. **Samo analiza: u `src-vba/` ništa nije
menjano.** Uz nju ide predlog tajmera za merenje na Windows-u.

| Fajl | Šta je |
|---|---|
| `README.md` (ovaj) | Odgovor, rang po očekivanoj dobiti, redosled, odluke koje su tvoje, šta nije provereno |
| `02-katalog-nalaza.md` | Svih 107 nalaza doslovno: lanac poziva, obrazac, učestalost, složenost, predlog, testovi i dokaz verifikatora |
| `03-tajmer.md` | Dizajn `modPerf`, 70 tačaka merenja vezanih za ID nalaza, protokol merenja i rizici |
| `modPerf-skica.bas.txt` | Skica modula tajmera (`.txt` da je ne pokupe VBA alati i hook) |

## 0. Šta ovo jeste, a šta nije

- **Stablo.** Analiza i svi brojevi linija važe za `main @f154530`. Grana je
  19.11. rebase-ovana na `main @1a2293d` (33 commita, uglavnom cenovnik i
  matični), pa sam proverio šta se pomerilo:
  - **Nepromenjeni** su `modPrint`, `modDokumenta`, `modAmbalaza`,
    `modAutoHladnjaca`, `modMasterSync`, `modJournaling`, `clsTransaction`,
    `modIzvestaj`, `modScrIzvestaji` i `modOtkup`. Zato OTK-1, HLD-1, knjiga
    ambalaže, članstvo otpremnice, AutoSave i Izveštaji važe sa **istim**
    linijama.
  - `modDataAccess`: samo `AuditUser` je postao `Public`.
  - `modScrDokumenti`: jedna linija komentara manje, pa su linije posle `2620`
    pomerene za −1.
  - `modOtkupUI` je pomeren za +9 do +340 linija. Na novom stablu:
    - `UiEvent` 4523;
    - `RefreshKpi` 4480;
    - `RefreshFromData` 6547;
    - `OsveziNavBrojace` 6513;
    - `FillZbirneCombo` 7220;
    - `RefillMaticniCombos` 7404;
    - `UiHover` 5540;
    - `RenderGrid` 3107;
    - `LayoutGrid` 2829;
    - `CommitDokument` 8821;
    - `SelectModeCore` 4154;
    - `UiChange` 5437;
    - `OnStanicaChanged` 9152;
    - `SaldoOMUkupno` 8346;
    - `RefreshOtpTraka` 956.
  - **UI-8 i dalje važi:** kvadratni sort zbirnih je netaknut u novom
    `FillZbirneCombo`, oko linije 7320.
  - **Novo na `main`-u, nije analizirano:**
    - `SelectModeCore` i `OnStanicaChanged` sada zovu `AutoFillCena`, a on
      zove `modCenovnik.TryVazecaCena` za jednu ili dve klase.
    - To je novi trošak pri promeni režima i stanice. Ako se potvrdi kaskada
      UI-7, ide i posle svakog snimanja.
    - Cena `TryVazecaCena` nije čitana. Tajmer je hvata kroz akciju
      `UI.Change.*`.

- **Ništa nije mereno.** Sva vremena u ovom dokumentu su **PROCENE** dobijene iz
  čitanja koda i pretpostavljenih veličina tabela. Sveska `.xlsm` nije u repou.
  Broj redova, formule u listovima i uslovno formatiranje se ne vide, a
  `run_vba` ne radi u Linux sesiji. Zato je rang ispod **hipoteza za merenje**,
  ne presuda. Tajmer (§5) postoji da je potvrdi ili obori.
- **Šta je provereno:**
  - Šest oblasti je analiziralo šest agenata, a svaki nalaz je zatim pročitao
    drugi agent kome je zadatak bio da ga obori. Od 71 nalaza: **49 potvrđeno,
    22 delimično potvrđeno, 0 oborenih**. Kod delimično potvrđenih verifikator je
    ispravio broj ili procenu, i ispravka je ušla u katalog.
  - Kritičar pokrivenosti je našao tri rupe, a tri dodatna agenta su ih
    istražila.
- **Šta NIJE nezavisno provereno:** 14 dodatnih nalaza verifikatora i 22 nalaza
  iz rupa. U katalogu su tako i označeni.
- **Sam sam proverio u kodu najvažnije tvrdnje i sve stoje:**
  - OTK-1: `modPrint.bas:705-742`;
  - HLD-1: `modDokumenta.bas:3092-3107` + `OtpremnicaJeIzdata` na `7650`;
  - UI-5: `modOtkupUI.bas:5265-5283`;
  - UI-8: `modOtkupUI.bas:7006-7013`;
  - IZV-7: `modOtkupBlok.bas:200-211`;
  - AutoSave i warm na 60 s: `modJournaling.bas:39-40, 820-825`,
    `modStornoWarm.bas:27`;
  - pogodak keša vraća kopiju: `modDataAccess.bas:136-138`;
  - nijedan `BeginTableCache` nije na putu upisa (grep).

## 1. Kratak odgovor

**Da, prostora ima, i to dosta.** Skoro sve se svodi na šest korenskih uzroka,
ne na 107 nezavisnih mesta:

1. **Kvadratne petlje nad celom tabelom** tamo gde treba jedan prolaz:
   - štampa otkupnog lista (OTK-1);
   - lista nevezanih otpremnica (HLD-1);
   - sort zbirnih u combo-u (UI-8);
   - RANG (IZV-7).

   Rastu sa kvadratom broja redova, pa u sezoni ovo dominira.
2. **Na putu upisa nema keša čitanja.** `BeginTableCache` se otvara samo na
   ekranima matičnih, Izveštaja i Storna. Svaki `GetTableData`, `LookupValue`
   i `FindRows` u transakciji zato iznova kopira celu tabelu iz lista. Iste
   velike tabele (`tblOtkup`, `tblAmbalaza`, `tblOtpremnica`) čitaju se 5–17
   puta po jednoj akciji.
3. **Ista provera se ponavlja nad nepromenjenim podacima:**
   - integritet knjige ambalaže, oko 10 puta po snimanju otkupa;
   - šema-kapija, 8–9 puta;
   - kanonsko članstvo otpremnice, po jednom za svakog člana bloka, što daje
     O(k·N), a pri vezivanju i ispravci izdate O(k²·N).
4. **Posle svake akcije se osvežava sve.** `RefreshFromData` briše ceo UI keš i,
   bez prozora keša, ponovo računa KPI (godišnji `ReportSaldoOM` da bi pročitao
   jednu ćeliju), značke svih ekrana, combo-e i mrežu. Na Izveštajima kroz isti
   put ide i **svaka promena konteksta** (tip, režim, entitet, opseg).
5. **Sinhroni izlaz na svakom snimanju:**
   - PDF otkupnog lista i otvaranje čitača (dugmad „Sačuvaj" i „Sačuvaj i
     štampaj" rade isto);
   - u malina režimu i PDF grupnog lista.
6. **Pozadinski poslovi okidaju u istom trenutku:** AutoSave i StornoWarm su oba
   na +60 s od poslednjeg commit-a. Ako operater tada kuca, forma stoji 2–12 s
   (PROCENA).

## 1a. Rang za stvarne veličine (dopuna 19.11.2026)

**Ulaz od vlasnika** (procena, ne izmereno):
- `tblOtkup` najviše oko 10.000 redova;
- `tblOtkupStavke` najviše 15.000, realno manje;
- `tblAmbalaza` najviše oko 50.000;
- `tblOtpremnica` i `tblZbirna` znatno manje od otkupa. **Pretpostavka ovde:
  1.000–3.000.**
- **Sve firme** koriste hladnjačku stanicu (RUTA_LANAC), **neke** malina režim.
- Office je najčešće 64-bitni.

**Štampa, kako sam razumeo:**
- **bez malina režima** pojedinačni otkupni list je obavezan uz svaki otkup;
- **u malina režimu** obavezan je grupni list, a pojedinačni otkupni listovi
  nisu potrebni.

**Provereno u kodu za ovu dopunu:**
- `ds` u `FillOtkupSablon` je **cela** `tblOtkupStavke`
  (`modPrint.bas:658`, `GetTableData(TBL_OTKUP_STAVKE)`).
- `Scr_Save` (`modScrDokumenti.bas:2174-2176`) zove `AutoLanacHladnjaca` pri
  svakom otkupu čija je ruta RUTA_LANAC. Lanac zatim ide:
  - `CreateOtpremnicaIzIzvora_TX` (`modAutoHladnjaca.bas:382`);
  - `modMasterSync.AutoZbirnaZaOtpremnicu` (`:392`), koja zove
    `modDokumenta.NevezaneOtpremnice()` (`modMasterSync.bas:1071`);
  - `NevezaneOtpremnice` za svaki nestorniran red `tblOtpremnica` zove
    `OtpremnicaJeIzdata` → `LookupValue`. To je puna kopija `tblOtpremnica` po
    redu (`modDataAccess.bas:631-635`), bez keša.
- `GRUPNI_OTKUP_PRINT_MODE` kad je prazan **prati `OTKUP_PRINT_MODE`**
  (`modConfig.bas:678`, `modPrint.bas:1196-1197`).

**Preračun.** Cene po jedinici su iz kataloga. Agenti se oko njih razilaze i do
10×, pa je za svaki red dat i donji raspon uz 10× jeftiniju jedinicu. Zaključak
mora da preživi i taj donji raspon.

| Rang | Šta | Ko plaća | Računica (PROCENA, nije mereno) | Na tvojim veličinama | Lek |
|---|---|---|---|---|---|
| **1** | **OTK-1** popuna otkupnog lista | svako snimanje otkupa, u firmama bez malina režima | ≈ N_stavke² iteracija × 0,3–1 µs | 3.000 stavki: 3–9 s; 10.000: 30–100 s; 15.000: 70–225 s. **Uz 10× jeftiniju iteraciju i dalje 3–10 s pri 10.000.** Snimanje se usporava kako sezona odmiče. | lokalna izmena u `FillOtkupSablon` (jedan prolaz skupi redove tog otkupa); test u `modTest`. Malina firme: konfig, v. §3 korak 1 |
| **2** | **HLD-1** `NevezaneOtpremnice` | **svako** snimanje otkupa na hladnjačkoj stanici (sve firme); izdavanje uz auto-zbirnu (`modScrDokumenti.bas:887`); lista NEVEZANE | ≈ N_otp × N_otp × 24 kopirane ćelije × 0,1–0,5 µs | 1.000 otpremnica: 2,4–12 s; 2.000: 10–48 s; 3.000: 22–108 s. **Uz 10× jeftiniju ćeliju: 0,2–1 s / 1–5 s / 2–11 s.** Raste kroz sezonu kao i #1. | lokalna izmena: indeks kolone statusa jednom, pa postojeći `IzdatoStatusJeIzdato(otp(i, cIzd))` umesto `LookupValue` po redu (`modDokumenta.bas:7660`, review #362); test `Test_ZBR_NevezaneSamoIzdateISlobodne` već meri ishod |
| **3** | **Knjiga ambalaže** (OTK-2 i spojeni) | svako snimanje sa ambalažom; u hladnjačkom lancu još i otpremnica i prijemnica istog klika (nije prebrojano) | 10 čitanja po otkupu × (kopija 50.000×16 ćelija + `KnjigaIntegritet` 2 prolaza po 2–100 µs po redu) | **od oko 1 s do desetina sekundi po snimanju otkupa.** Raspon je 50× jer agenti različito cene integritet po redu. | **odluka D1.** Pri 50.000 redova integritet verovatno dominira, pa memo (OTK-2) vredi znatno više od prozora keša. Prvo izmeriti udeo `KnjigaIntegritet`, pa memo kroz `pre-flight` + `dokaz.py --izbor` |
| 4 | **Ostatak hladnjačkog lanca**: HLD-2..5 i revalidacija u `CreateOtpremnicaIzIzvora_TX` | svako snimanje otkupa na hladnjači | 9 kopija `tblOtkup` (10.000×29) pre prvog upisa; 3 prolaza strogog čitača stavki; 4 TX sa snapshotima, `tblAmbalaza` (800.000 ćelija) još 2 puta | 0,5–3 s po kliku | prozor keša u glavi lanca (HLD-2), zatim D2 |
| 5 | **Posle svakog snimanja** (UI-7, pa UI-1 i spojeni) | svako snimanje i radnja | `tblOtkup` 290.000 ćelija × 7–9 čitanja + KPI (`ReportSaldoOM` za godinu) + značke | 0,3–3 s po snimanju | prvo izmeriti UI-7, pa garda, pa prozor oko `RefreshFromData` |
| 6 | **UI-8** sort zbirnih u combo-u | svako snimanje. U lancu svaki otkup pravi i zbirnu, pa N_zbr raste kao N_otp. | ≈ N_zbr²/2 poređenja i zamena × 0,1–0,5 µs | 1.000: 0,05–0,25 s; 3.000: 0,5–2 s | lokalna (`QuickSortIndex` postoji). **Pažnja:** zatečeni sort nije stabilan, pa redosled zbirnih sa istim datumom treba ugovoriti testom |
| 7 | **AutoSave i warm** u istoj sekundi | jednom po pauzi | sveska raste sa 50.000 redova knjige | 2–12 s zamrznute forme | odluka D4 |
| 8 | **Izveštaji**: IZV-1 (kontekst ide kroz pun refresh); **IZV-4** zbirna ambalaža | svaki klik konteksta; zbirni prikaz ambalaže | IZV-4: broj naloga × cela knjiga (50.000) sa integritetom i šemom | IZV-1: 0,3–2 s po kliku. **IZV-4: od nekoliko sekundi do više desetina sekundi** za zbirni prikaz, zavisno od broja naloga (nije dat) i od cene integriteta | IZV-1 lokalna; IZV-4 ide uz D1 |
| 9 | Hover i pretraga na Dokumentima (UI-5, UI-9) | svaki prelaz miša i svaki otkucaj | v. §2 | 20–150 ms i 50–500 ms | lokalna |

**Šta se promenilo u odnosu na §2:**
- **HLD-1 je skočio sa 9. na 2. mesto.** Hladnjačku stanicu koriste sve firme,
  pa kvadratni prolaz plaća svako snimanje otkupa, ne retki slučaj.
- **Knjiga ambalaže je ojačala:** 50.000 redova je pet puta više od pretpostavke
  sa kojom su agenti računali.
- **PDF bez koda pomaže samo malina firmama.** Firmama bez malina režima
  pojedinačni otkupni list ostaje, pa im OTK-1 mora da se ispravi u kodu.

**Tajmer na 64-bitnom Office-u:** `Currency` trik i `PtrSafe` deklaracija rade.
`LongLong` bi takođe radio, ali skica ostaje na `Currency`, da isti kod radi i
na retkim 32-bitnim mašinama. Test `T_Perf_QpcMonoton` dovoljno je pustiti na
64-bitnom.

## 2. Rang po očekivanoj dobiti (opšti, pre podataka o veličinama; za tvoje veličine važi §1a)

Rang je moja procena. Ulazi su učestalost akcije, rast sa veličinom podataka i
cena po pozivu iz kataloga. **Konačan redosled zavisi od stvarnih veličina
tabela (§8).**

Legenda za kolonu „Vrsta izmene":
- **konfig** — bez koda;
- **lokalna** — jedna procedura, bez promene ugovora;
- **odluka** — traži tvoj izbor ili `pre-flight`.

| # | Šta | Tok i učestalost | Nalazi (spojeni duplikati) | Rast | PROCENA dobiti (nije mereno) | Vrsta izmene | Status |
|---|---|---|---|---|---|---|---|
| 1 | Popuna otkupnog lista: dvostruka petlja `For rb = 1 To UBound(ds,1)` × ceo `tblOtkupStavke` | svako snimanje otkupa (osim kad je štampa OFF) | OTK-1 | **N²** | 1.000 stavki: 0,3–1 s; 3.000: 3–9 s; 10.000: 30–100 s **po snimanju** → skoro 0 | lokalna (`FillOtkupSablon`) | potvrđeno + moja provera |
| 2 | Posle svakog snimanja: pun `RefreshFromData` bez prozora keša | svako snimanje, izmena, radnja nad redom | UI-1 = ZAJ-08 = IZV-8; UI-2, UI-3, UI-6, UI-11, UI-12, UI-D1, IZV-D1, ZAJ-D2 (značka BANKA_NALOZI), DOK-UI1, DOK-D1 | linearno, ×4–12 | 0,2–2 s po snimanju; ZAJ-D2 sam 0,3–10 s, zavisno od broja kooperanata | lokalna, ali **posle #3** | uglavnom potvrđeno; ZAJ-D2 nije nezavisno |
| 3 | Kaskada Change pri ponovnom punjenju combo-a (`RefillMaticniCombos` bez `mLoading`) | svako snimanje | UI-7; UI-D2 (`BulkPushPendingForStanica` čita `tblOtkup` ćeliju po ćeliju i piše mimo `modDataAccess`) | — | 0,5–3 s po snimanju uz cloud, inače 0,1–0,5 s | lokalna, prvo **izmeriti** (ne zna se da li MSForms diže Change na `.Clear`) | delimično |
| 4 | Knjiga ambalaže: čitanje + `KnjigaIntegritet` (2 prolaza) + šema-kapija, ponovljeni nad nepromenjenom knjigom | svako snimanje otkupa ili dokumenta sa ambalažom; storno i ispravka | OTK-2, AMB-1, DOK-D3, ZAJ-01, ZAJ-D1, OTK-8 = ZAJ-10 = SCH-1, OTK-14, AMB-ISP-1, AMB-STO-2, AMB-ODL-3, AMB-INT-4, AMB-DUP-5, GRP-3; u Izveštajima IZV-4 | linearno, ×10–16 | 0,2–1 s po snimanju na svakih 1.000 redova knjige; storno ili ispravka u lancu 1–6 s pri 10.000 redova | **odluka** (§4, D1); kritična invarijanta → `dokaz.py --izbor` obavezan | potvrđeno ili delimično; AMB-* iz rupe nisu nezavisno |
| 5 | Sinhroni PDF i otvaranje čitača | svako snimanje otkupa; prijemnica u malina režimu | OTK-3, GRP-1, GRP-2, GRP-6; uz to OTK-16, GRP-4, GRP-5, GRP-7 | konstanta | 0,5–3 s po snimanju otkupa, 0,5–4 s po prijemnici | **konfig** (§3, korak 1) | OTK-3 potvrđeno; GRP-* nisu nezavisno |
| 6 | Revalidacija članstva otpremnice po članu, plan lanca više puta | izdavanje, izmena nacrta, vezivanje bloka (i automatski posle svakog otkupa uz aktivan nacrt), ispravka izdate, potezi i storno u lancu | OTP-1 = IZM-1, OTP-2, IZM-3, IZM-D1; OTP-3 = IZM-9 = IZM-D2 = OTK-5; IZM-2; OTK-13, IZM-4, IZM-5, IZM-6, IZM-8; LNC-PLAN-6, LNC-CLAN-7, LNC-PRE-8; OTK-4 = DOK-D2 | **k·N**, kumulativno **k²·N** | za k = 20 i 10.000 otkupa sekunde po potezu (procene u katalogu) | delom lokalna (prozor oko read-only kapija), delom **odluka** (§4, D2) | uglavnom potvrđeno; LNC-* nisu nezavisno |
| 7 | AutoSave i StornoWarm u istoj sekundi, 60 s posle commit-a | jednom po pauzi u radu | OTK-12 = OTK-15 = ZAJ-03 | veličina sveske | 2–12 s zamrznute forme po okidanju | **odluka** (pomeriti jedan od dva, §4, D4) | potvrđeno, uz nijansu kritičara |
| 8 | Izveštaji: promena konteksta ide kroz pun `RefreshFromData` i briše sve snimke | svaki klik tip/režim/entitet/opseg | IZV-1 = UI-4; IZV-2, IZV-9 | — | 0,3–2 s po promeni, plus ceo `Report*` pri povratku na već viđen kontekst | lokalna | potvrđeno |
| 9 | `NevezaneOtpremnice`: `LookupValue` nad `tblOtpremnica` za svaki red `tblOtpremnica` | svako snimanje otkupa na hladnjačkoj stanici; lista NEVEZANE; malina auto-zbirna (`modMasterSync`) | HLD-1, HLD-1b | **N²** | 300 otpremnica: 0,2–1 s; 1.000: 2,4–12 s; 3.000: 22–108 s | lokalna: koristi postojeći `IzdatoStatusJeIzdato` nad ćelijom reda (komentar na `modDokumenta.bas:7660`, review #362, traži baš to) | nalaz iz rupe + moja provera koda |
| 10 | Izveštaji, unutar `Report*` | zbirni oblici | IZV-3 (saldo po stanici gradi strukture nezavisne od stanice), IZV-5 (kupci), IZV-6, IZV-10, IZV-11, IZV-7 (bubble sort, O(K²)) | S·N, K·N, K² | stotine ms do sekundi po zbirnom prikazu | lokalna; IZV-3 i IZV-5 traže **odluku** D3 | potvrđeno ili delimično |
| 11 | Dokumenti: hover i pretraga | svaki prelaz miša preko reda; svaki otkucaj | UI-5 (pun `RenderGrid` 22×14 na hover), UI-9 (sve izvedene mape iznova po otkucaju), UI-10 | — | 20–150 ms po hover-u; 50–500 ms po otkucaju | lokalna | potvrđeno ili delimično |
| 12 | Combo zbirnih: zamenski sort cele istorije | svako snimanje | UI-8 | **N²** | 500 zbirnih: 10–40 ms; 3.000: 0,5–2 s | lokalna (postoji `QuickSortIndex` / `QSortIdx`) | potvrđeno + moja provera |
| 13 | Hladnjački lanac, ostalo | snimanje otkupa na hladnjači | HLD-2 (9 kopija `tblOtkup` pre prvog upisa), HLD-3, HLD-4, HLD-5 (4 TX umesto 1) | linearno | 0,3–2 s po kliku | lokalna (prozor u glavi lanca) | nalazi iz rupe |
| 14 | Sitno u zajedničkom sloju i u piscima | sve | OTK-6, OTK-7 = DOK-ZAJ1, OTK-9, OTK-10, OTK-11 = ZAJ-02 = PRJ-2, ZAJ-04, ZAJ-05, ZAJ-06 = IZM-12, ZAJ-07, ZAJ-09, NOV-1, PRJ-1, ZBR-1 = IZM-10, PAL-1, OTP-4, IZM-7, IZM-11 | linearno ili konstanta | 10–150 ms svaki | uglavnom lokalna | potvrđeno |

**Važna ograda za sve predloge sa prozorom keša (IZM-D3):** pogodak u kešu i
dalje vraća **kopiju** celog niza (`GetTableData = mTableCache(tblName)`,
`modDataAccess.bas:137`).
- Prozor zato štedi COM čitanje iz lista.
- Ne štedi memorijsku kopiju niti validaciju koja ide posle (`KnjigaIntegritet`,
  `AktivnoOtpClanstvoPoKanonu`).
- Procene ušteda u OTP-1/2/3, IZM-1..10, PRJ-1, ZBR-1 i UI-1 su zbog toga
  **precenjene**.
- Tačan odnos daje tek merenje sa brojačima iz §5.

## 3. Predloženi redosled

**Urađeno 19.11.2026** (grana `claude/funny-turing-xc1u31`, svaki rez svoj
commit):

| Rez | Šta | Test | Status |
|---|---|---|---|
| IZV-1 | Promena tipa, režima, entiteta i opsega na Izveštajima osvežava samo prikaz (`modOtkupUI.OsveziPrikazEkrana`), bez punog `RefreshFromData`. Povratak na viđen kontekst ide iz keša. | `modTest` 384 `T_Izv_KontekstOsvezavaPrikazNePodatke` (klik kroz pravu ljusku) | **zeleno 19.11.** (FULL + dokaz) |
| OTK-1 | Popuna otkupnog lista bira stavke jednim prolazom (`modPrint.OtkupListRedoviStavki`) umesto kvadratne petlje. Jednakost sa starim algoritmom je proverena Python simulacijom nad 20.000 tabela. | `modTest` 385 `T_OtkList_RedoviStavkiKanonskimRedom` (čista funkcija) | **zeleno 19.11.** (FULL + dokaz) |
| HLD-1 | `NevezaneOtpremnice` čita status iz svog reda umesto `LookupValue` po redu. Sidro sabotaže `zbirna-nevezane-nude-nacrt` je pomereno na novu liniju. | BFP `Test_ZBR_NevezaneSamoIzdateISlobodne` | **zeleno 19.11.** (FULL + dokaz) |
| IZV-3, IZV-5 | Keš izvedenih nizova u prozoru čitanja (`modDataAccess.IzvedenoTryGet`/`IzvedenoPut`): `StavkeOtkupaRedovi`, `DoplataStavkeRedovi` i `StavkePrijemniceRedovi` se u jednom zbirnom prikazu grade jednom, a ne po stanici ili kupcu. Upis prazni keš; rečnici se ne keširaju (D3). | `modTest` 386 `T_KesIzvedeno_ZbirniIstiUzKes` (isti sadržaj sa i bez keša) | **zeleno 19.11.** (FULL + dokaz) |
| IZV-2 | Naziv entiteta se čita iz šifarnika samo kad se prikazani kontekst promeni, ne na svaki otkucaj pretrage i klik čipa. Pomerena sidra `izvestaji-pretraga-puni-iznova` i `izvestaji-rang-kontekst-prazan`. | `modTest` 387 `T_Izv_NazivEntitetaJednomPoKontekstu` | **zeleno 19.11.** (FULL + dokaz) |
| IZV-9 | Dinamički čipovi vrste i sorte čitaju keš ljuske (`modUiData.CachedTable`) umesto tabela sa lista. | `modTest` 388 `T_Izv_CipoviCitajuKesLjuske` | **zeleno 19.11.** (FULL + dokaz) |
| IZV-7 | Rang kooperanata: stabilno spajanje po broju (`modArrayUtils.IndeksiPoBrojuOpadajuce`) umesto mehurastog sortiranja, isti poredak (Python simulacija, 20.000 nizova). | `modTest` 389 `T_Rang_StabilnoOpadajuce` | **zeleno 19.11.** (FULL + dokaz) |
| PERF-REV-02 (pregled 19.11.) | `NevezaneOtpremnice` primenjuje pravilo jedinstvenosti iz `StavkeOtpremniceRedovi` (greška 1941): dupli `OtpremnicaID` pada po imenu, umesto da se bira red po slučajnom redosledu. | `modTest` 390 `T_Nevezane_DupliIdPadaPoImenu` | **zeleno 19.11.** (FULL + dokaz) |
| PERF-REV-03 (pregled 19.11.) | `modDataAccess.KesTabeleZastareo`: rollback (`RestoreTable`) i oznaka sinhronizacije (`BulkPushPendingForStanica`) obaraju keš čitanja, pa ugovor ne zavisi od pretpostavke da u prozoru niko ne piše. | `modTest` 391 `T_KesTabele_RollbackZastareva` | **zeleno 19.11.** (FULL + dokaz) |
| D1-IZV (20.11., posle merenja §3a) | Dokaz integriteta knjige ambalaže se u **prozoru čitanja** pamti posle prvog uspeha (`modAmbalaza.KnjigaZaCitanje` kroz `IzvedenoTryGet`/`IzvedenoPut`). Svaki upis, rollback i `KesTabeleZastareo` ga brišu, a pamti se samo uspeh. Van prozora, dakle i u svakom upisu, dokaz ide kao pre. Audit (`AmbKnjigaNalazi`) dokazuje uvek iznova. Pre koda je prošao `pre-flight`. | `modTest` 392 `T_KesKnjiga_DokazJednomUProzoru`; sabotaže `kes-knjiga-dokaz-gluv`, `kes-knjiga-upis-ne-obara-dokaz` | **zeleno 20.11.** (FULL + dokaz `kes-knjiga` 2/2) |
| D1-UPIS (20.11.) | Van prozora čitanja `KnjigaZaCitanje` pamti **dokazan prefiks** knjige: ćelije u 12 kolona koje dokaz čita, raspored kolona i rečnik vrsta. Sledeće čitanje poredi prefiks ćeliju po ćeliju (isti tip, tekst binarno), a **istu** funkciju `KnjigaIntegritet`, pokrenutu od prvog novog reda, prolaze samo dodati redovi. Promenjen, skraćen ili premešten prefiks, druge kolone ili greška u ćeliji znače pun dokaz. Ne zavisi od toga ko je pisao (upis mimo `modDataAccess`, ručna izmena, rollback). Pre koda je prošao `pre-flight`; dokaz je čista funkcija 12 kolona knjige (`AmbPrenosStrukturaProblem` ne čita tabele). | `modTest` 393 `T_KnjigaDokaz_NastavakIstiSud`; sabotaže `knjiga-nastavak-gluv`, `knjiga-nastavak-ne-poredi-dokazano`, `knjiga-nastavak-ne-dokazuje-nove` | **zeleno 20.11.** (FULL + dokaz `knjiga-nastavak` 3/3) |
| UI-7 (20.11., posle merenja §3a) | `RefillMaticniCombos` puni liste konteksta (OM, vozač, vrsta, sorta, tip ambalaže) pod `mMaticniFill`, pa njihov programski `Change` ne pokreće handlere. Izbor koji se punjenjem **stvarno** promeni (npr. deaktivirano otkupno mesto) ide kroz handler jednom, posle punjenja. Otvoreni blokovi se posle upisa osvežavaju eksplicitno (do sada ih je osvežavala usputna kaskada). `RefreshFromData` ne prlja čistu formu: rešen je i terenski simptom „Nesačuvan unos” posle snimanja i Sync-a. | `modTest` 394 `T_Refresh_MaticneListeBezKaskade`; sabotaže `ui7-liste-bez-garda`, `ui7-nestao-izbor-bez-handlera`, `ui7-osvezavanje-prlja-formu` | **zeleno 20.11.** (FULL + dokaz `ui7` 3/3) |
| UI-1 (20.11.) | `RefreshFromData` radi u prozoru čitanja, posle `ResetCache` do kraja. Keš ostaje ispravan i kad se usput piše: upis kroz `modDataAccess`, rollback i `KesTabeleZastareo` ga brišu, a kaskada je pod gardom UI-7. Test seam `CitanjaSaListaTest*` broji čitanja sa lista po tabeli. | `modTest` 395 `T_Refresh_CitaTabeluJednom`; sabotaža `ui1-osvezavanje-bez-prozora` | **zeleno 20.11.** (FULL + dokaz `ui1` 1/1) |

- **Ovde je zeleno:** sve statičke provere iz CI-ja (`vba_check`, `sabotaza`,
  `who_writes`, `gen_schema_module`, kapije self-update-a, `vba_gate`).
- **Pušteno na Windows-u 19.11.** (operater, izbor testova): 384, 385, 386 i
  BFP `Test_ZBR_NevezaneSamoIzdateISlobodne` zeleni;
  `dokaz.py --izbor zbirna-nevezane` 2/2 DOKAZANO.
- **FULL 19.11.** nad stanjem pre ojačanja po pregledu (389 testova): svih 12
  suite-ova zeleno, `RunAllTests` 389/389, BFP 3.499/3.499, GREEN marker upisan.
  Testovi 387–389 i dokaz `izvestaji-pretraga-puni-iznova` i
  `izvestaji-rang-kontekst-prazan` 2/2 DOKAZANO. Compile je u prolazu preskočen.
- **Posle ojačanja 19.11.:** testovi 387–391 zeleni (izbor), dokaz dva pomerena
  sidra Izveštaja ponovljen 2/2.
- **Dokaz novih testova 19.11.:** `dokaz.py --izbor` nad 10 novih sabotaža
  (384–391) 10/10 DOKAZANO.
- **FULL posle ojačanja 19.11.:** svih 12 suite-ova zeleno, BFP 3.499/3.499,
  GREEN marker upisan nad konačnim `src-vba/` grane.
- **Pregled 19.11. (TEST390-REV-01):** testovi 386, 387 i 390 su pisali u fixture
  kroz `UpdateCell`, koji pečatira `ModifiedAt`/`ModifiedBy`. Vraćanje vrednosti
  nije vraćalo audit. Sada kvar ide u transakciji sa snimkom tabele i obaveznim
  rollback-om, a test tvrdi da je cela tabela, i audit, posle testa ista kao pre.
  Pušteno 19.11.: 386, 387, 390 zeleni (tabele posle testa iste, i audit, pa je
  potvrđena i vernost rollback-a); dokaz `izv3-`, `izv2-naziv`, `hld1-` 4/4.
- **Tajmer 19.11.** (grana `claude/merenje-modperf`): testovi 392–398 (od 20.11.
  396–402: 392 D1-IZV, 393 D1-UPIS, 394–395 UI-7/UI-1) zeleni,
  dokaz `perf-` 7/7. Test 396 potvrđuje da merač ne dira `Err` pozivaoca.
- **Ostaje pred merge** (ručno, nad kopijom prave sveske):
  - pun prolaz `python tools/run_vba.py` posle izmene testova 386, 387, 390;
  - ručni Compile (`run_vba` ga preskače uz suite);
  - poređenje štampe otkupnog lista pre i posle za isti dokument (obe klase,
    cena, bruto/neto, ambalaža);
  - hladnjački lanac i lista NEVEZANE nad pravim podacima;
  - ~~merenje vremena pre i posle~~ — **urađeno 20.11.**, rezultat je u §3a.
  - za D1-IZV: ~~test 392, dokaz `kes-knjiga` 2/2 i pun prolaz~~ (zeleno
    20.11.: 12 suite-ova, `RunAllTests` 392/392, BFP 3.499/3.499, GREEN
    marker); ~~ponovno merenje zbirnog prikaza~~ (20.11.: 4,41 → 1,80 s).
  - za D1-UPIS: ~~test 393, dokaz `knjiga-nastavak` 3/3 i pun prolaz~~ (zeleno
    20.11.: 12 suite-ova, `RunAllTests` 393/393, BFP 3.499/3.499, GREEN marker);
    ~~merenje snimanja otkupa~~ (20.11.: prosek ≈5,6 s, bez ijedne pune
    provere knjige).
  - za UI-7 i UI-1: ~~testovi 394 i 395, dokaz `ui7 ui1` 4/4 i pun prolaz~~
    (zeleno 20.11.: 12 suite-ova, `RunAllTests` 395/395, BFP 3.499/3.499,
    GREEN marker); ~~merenje snimanja~~ (20.11.: osvežavanje 2,25 → 1,46 s);
    ~~ručno~~ (operater 20.11.: posle snimanja nema „Nesačuvan unos”, isplaćen
    blok se više ne nudi, deaktiviran izabran OM se prazni iz izbora).
    - Usput, nije formalno merenje: BFP je na istoj mašini trajao 483 s umesto
      629 s posle D1-IZV (−23%), a `RunAllTests` 241 s umesto 259 s. BFP
      stalno knjiži ambalažu, pa je to prvi znak da nastavak dokaza radi u
      putu upisa.

**Sledeće za Izveštaje** (ne zavise od D1):
- ostatak IZV-3: po stanici se i dalje grade rečnici kooperanata i
  čita knjiga ambalaže; to ide uz D1 (knjiga) i D3 (rečnici).

Zbirna ambalaža (IZV-4) ide uz D1.

1. **Bez ijedne linije koda, odmah** (#5):
   - **Podešavanja → „Otvori PDF posle izrade"** (`PDF_OTVORI_POSLE`) = NE
     (`modPodesavanja.bas:125`). PDF se i dalje pravi i arhivira, samo se ne
     otvara i ne uzima fokus. Važi za sve firme kojima ne treba da se čitač
     otvara posle svakog snimanja.
   - **Firme u malina režimu**, kojima pojedinačni otkupni list ne treba (§1a).
     Redosled je bitan:
     1. **prvo** `GRUPNI_OTKUP_PRINT_MODE` eksplicitno na `PDF` ili `PRINT`;
     2. **tek onda** `OTKUP_PRINT_MODE` = `OFF`.

     Grupni kad je prazan prati otkupni (`modPrint.bas:1196-1197`), pa bi
     obrnut redosled ugasio i grupni list. Ovim se skidaju OTK-1, OTK-3,
     OTK-14 i OTK-16 sa svakog snimanja otkupa.
     **Cena:** `OFF` gasi i **ručnu ponovnu štampu** otkupnog lista iz
     Izveštaja. Lanac je `modScrIzvestaji.bas:2956, 2963, 3012` →
     `ReprintOtkupniListByOtkupID` (`modPrint.bas:518`) → `OutputOtkupniList`,
     koji kod `OFF` ne radi ništa (`modPrint.bas:483-492`).
     - Ako i malina firmi ponekad treba pojedinačni list, ovaj prekidač nije
       besplatan, pa je bolji put ispravka OTK-1 u kodu.
     - Grep ne nalazi drugog korisnika `OTKUP_PRINT_MODE` osim ova tri: otkupni
       list, ponovna štampa i grupni kad je prazan.
   - **Firme bez malina režima:** otkupni list ostaje. Kod njih bez koda pomaže
     samo `PDF_OTVORI_POSLE`, a OTK-1 ide u korak 3 kao prvi rez.
   - **Pažnja:** grupni list nosi klauzulu PDV nadoknade (čl. 34 ZPDV,
     `modPrint.bas:1385-1392`). Da li je papir obavezan je tvoja poslovna i
     pravna odluka, ne tehnička.
2. **Tajmer na zasebnoj grani merenja** (§5), zatim jedan krug protokola nad
   kopijom prave sveske. Brojači (koliko puta je `tblAmbalaza` pročitana po
   snimanju) su deterministički i potvrđuju ili obaraju nalaze nezavisno od
   šuma.
3. **Lokalne algoritamske ispravke, bez promene ugovora,** svaka u svom rezu sa
   testom koji meri ponašanje:
   - #1 OTK-1;
   - #9 HLD-1 (po §1a drugi po dobiti, odmah posle OTK-1);
   - #12 UI-8;
   - IZV-7;
   - #11 UI-5 (hover boji samo stari i novi red; `PaintRow` postoji);
   - #8 IZV-1 (promena konteksta Izveštaja ne ide kroz pun `RefreshFromData`).

   Ovde je najveći odnos dobiti i rizika.
4. **Posle merenja:**
   - #3 UI-7 (garda kaskade), pa tek onda #2 (prozor keša oko
     `RefreshFromData`). Redosled je bitan: kaskada `OnStanicaChanged` →
     `BulkPushPendingForStanica` piše `tblOtkup` mimo `modDataAccess`, pa bi
     prozor otvoren pre garde čitao zastareo keš.
   - Zatim #10, #11 (UI-9) i #13.
5. **Uz odluke iz §4 i skill `pre-flight`:** #4 (knjiga ambalaže), #6 (članstvo
   otpremnice), #7 (raspored AutoSave i warm).

Za svaku izmenu ponašanja važi:
- test u `modTest`;
- `python tools/vba_check.py`;
- ciljana suite na Windows-u;
- kod kritičnih invarijanti (#4, #6) još i
  `python tools/dokaz.py --izbor --radnika 4 <prefiks>`.

Iz web sesije se ništa od toga ne može prijaviti kao zeleno.

## 3a. Merenje pre i posle (20.11.2026)

**Uslovi:**
- Dve sveže kopije prave sveske, iz iste osnove; u obe su cloud sync i
  monitoring ugašeni.
- Kod:
  - „pre" = `claude/merenje-modperf-pre` @7d55949 (`main @1a2293d` + merač);
  - „posle" = `claude/merenje-modperf` @5313843 (ova grana + isti merač).

  Obe stoje na istom `main`-u, pa je razlika samo u rezovima iz §3.
- `modPerf`:
  - nivo 1, prag 0 ms, QPC 10 MHz;
  - cena merača oko 9 µs po sekciji i 3 µs po brojaču.
- Veličine tabela u merenju:
  - `tblOtkup` i `tblOtkupStavke` po 1.813;
  - `tblOtpremnica` 976, `tblZbirna` 977, `tblPrijemnica` 949;
  - `tblAmbalaza` 5.239, `tblKooperanti` 104.
- Scenario, isti redosled i isti unosi u obe kopije:
  - Izveštaji: saldo zbirno i pojedinačno, izbor OM-a, kartice tipa, liste i
    čipovi;
  - zatim 5 snimanja otkupa na običnoj i 5 na hladnjačkoj stanici;
  - na kraju mirovanje (AutoSave).
- META kontrola (`TableCacheDepth`, `Nezatvorenih`, `Neuparenih`, `Preliv`) je 0
  u obe sesije.
- Vremena su **aktivna**, bez čekanja na dijalog. Ponavljanja je malo:
  - Izveštaji: 1 hladno + 1–2 topla po radnji;
  - snimanje: 1 + 9.
- Izvor brojeva su sažeci sesija (`Perf_sazetak_2026-11-20_072031` i
  `_080348`). Kasnija kontrolna snimanja nisu uračunata.

**Rezultat:**

| Radnja | Pre | Posle | Rez |
|---|---|---|---|
| **Snimanje otkupa**, prosek 9 toplih | 12,5 s | 10,5 s | −2,0 s (−16%) |
| PDF otkupnog lista, po snimanju | 2,27 s | 1,27 s | OTK-1, −44% |
| Nevezane otpremnice, po hladnjačkom snimanju | 2,02 s | 0,03 s | HLD-1, −98% |
| Hladnjački lanac ukupno, po hladnjačkom snimanju | 6,5 s | 4,5 s | HLD-1 |
| Čitanja `tblOtpremnica` u 10 snimanja | 4.940 (115 mil. ćelija) | 325 (7,6 mil.) | HLD-1, −93% |
| **Izveštaji: Zbirno saldo**, ponovni klik | 5,8 s | 0,06–0,18 s | IZV-1 |
| Zbirno saldo, prvi klik | 5,7 s | 4,2 s | snimak se računa jednom |
| Pojedinačno | 1,9 s | 0,06–0,6 s | IZV-1 |
| Tip OM / Kupci | 5,8 / 1,4 s | 0,06 / 0,13 s | IZV-1 (v. ispod) |
| Izbor iz liste | 0,19 s, prvi 1,9 s | 0,09 s | IZV-1 |

- **Snimanje.** Očekivano je ≈1,0 s na PDF-u svakog snimanja i ≈2,0 s na
  nevezanim otpremnicama svakog hladnjačkog snimanja, u proseku ≈2,0 s.
  Izmereno je 2,02 s.
- **Kontrola da je scenario isti.** Brojači čitanja tabela koje nijedan rez ne
  dira su u snimanju identični pre i posle (`tblOtkup` 509, `tblAmbalaza` 155,
  `tblPrijemnica` 370, `tblKooperanti` 1.963, `tblSEFConfig` 691). Razlika je
  dakle samo u rezovima.
- **Izveštaji.** Klik na kontekst više ne ide kroz pun `RefreshFromData`. Pre je
  svaki klik čitao `tblKooperanti` oko 178 puta, `tblPrijemnica` 34 i `tblOtkup`
  14–22 puta sa lista, i nosio oko 1,25 s van samog snimka. Posle su to 0–2
  čitanja.
- **Prema proceni iz §1a:**
  - HLD-1 je za 1.000 otpremnica procenjen na 2,4–12 s, u donjem rasponu
    0,2–1 s. Izmereno pri 976: 2,02 s pre, 0,03 s posle.
  - OTK-1: pri 1.813 stavki ceo otkupni list (popuna i PDF) je pao sa 2,27 na
    1,27 s. Procena 3–9 s je bila za 3.000 stavki. Uklonjeni deo raste sa
    kvadratom broja stavki, pa dobit raste kroz sezonu.

**Nije izmereno ili nije uporedivo:**
- **IZV-7 (RANG):** izmereno samo posle (113 ms); pre nema broja.
- **IZV-9 (čipovi):** 34–40 ms i pre i posle. Pri ovoj veličini dobit nije
  merljiva.
- **IZV-3 / IZV-5:** prvi zbirni snimak je pao sa 4,4 na 4,1 s (SALDO) i sa 4,4
  na 4,2 s (ROBA). Dobit je mala. Ostatak je računanje, ne čitanje lista: posle
  su skoro sva čitanja u snimku pogoci keša. Za raščlanjavanje treba merenje na
  nivou `DETALJ`.
- **Tip OM 5,8 → 0,06 s:** posle je snimak već bio izračunat ranije u sesiji, a
  IZV-1 ga više ne briše. Broj meri povratak na viđen kontekst, ne prvo
  računanje.
- **Kartica Kooperanti nije uporediva:** posle je računat drugi snimak
  (KARTICA P, 0,54 s).

**Šta i dalje nosi vreme:**

| Gde | Trajanje (posle) | Napomena |
|---|---|---|
| Ostatak hladnjačkog lanca | 4,5 s po hladnjačkom snimanju | self `OTK.AutoLanacHladnjaca`, bez nevezanih otpremnica |
| Upis (`ScrSave` van izmerenih sekcija) | ≈3,7 s po snimanju | neraščlanjeno |
| Osvežavanje posle snimanja | 2,8 s po snimanju | pun `RefreshFromData`, opravdan jer su se podaci promenili |
| PDF otkupnog lista | 1,27 s po snimanju | |
| Programska promena OM-a (UI-7) | 0,8 s po snimanju | 3 `Change` po snimanju, ≈263 ms svaki |
| Prvi zbirni snimak (SALDO ili ROBA Z) | 4,1–4,2 s | računanje, v. gore |
| AutoSave | 7,3–7,5 s | forma stoji dok traje |

- Čitanja koja nijedan rez ne dira, po snimanju:
  - `tblOtkup` 51 put (2,7 mil. ćelija);
  - `tblAmbalaza` 15,5 puta (1,3 mil.);
  - `tblKooperanti` 196, `tblSEFConfig` 69.
- Udeo knjige ambalaže (D1) nije izdvojen, za to treba nivo `DETALJ`.

**Raščlanjavanje zbirnog snimka** (drugo merenje 20.11., sekcije na mernoj
grani, jedan klik):

| Klik | Ukupno | Dokaz integriteta knjige | Ostalo veće |
|---|---|---|---|
| Zbirno saldo (OM) | 4,41 s | 9× ≈ 325 ms = **2,93 s** (66%) | saldo-prolaz knjige po stanici 0,76 s; otkup 0,37 s (zbir stavki 0,13 s); šema-kapija knjige 9× 0,14 s |
| Zbirna roba (OM) | 4,32 s | 9× = **2,86 s** | isto |
| Snimanje otkupa (hladnjača) | 10,4 s | **15× = 4,82 s** (46%) | šema-kapija 12× 0,19 s |

- Zbirni saldo zove pojedinačni saldo po stanici, a svaki ponovo dokazuje istu,
  nepromenjenu knjigu (5.246 redova). Lek je D1-IZV iz §3.
- **Izmereno posle D1-IZV (20.11., treće merenje, prvi klik posle otvaranja):**

  | Klik | Pre reza | Posle reza |
  |---|---|---|
  | Zbirno saldo (OM) | 4,41 s | **1,80 s** (−59%) |
  | Zbirna roba (OM) | 4,32 s | **1,86 s** (−57%) |
  | Dokaz knjige u zbirnom kliku | 9× = 2,93 s | 1× = 0,31 s; 8 iz keša (`AMB.KnjigaDokaz.KES`) |

  Jedini preostali dokaz je u `IzvStaniceIzPodataka` (univerzum stanica),
  pre petlje po stanicama.
- **Šta i dalje nosi zbirni klik (1,8 s):**
  - saldo-prolaz knjige po stanici, 8× ≈ 95 ms = 0,76 s;
  - otkup po stanici 0,34 s (zbir stavki 0,12 s);
  - jedan dokaz knjige 0,31 s;
  - šema-kapija knjige 9× = 0,14 s.

  Sledeći rez bi bio jedan prolaz kroz knjigu za sve stanice. To je nov
  čitalac knjige, pa ide kroz `pre-flight`.
- **Usput izmereno:** Roba pojedinačno (jedan OM) traje 1,57 s, bez
  raščlanjavanja (`ReportOtkupRoba` nema sekcije).

**Snimanje otkupa posle D1-UPIS** (četvrto merenje 20.11.: 4 snimanja, 2
hladnjačka i 2 obična, isti odnos kao u §3a):

| | Pre (`main`) | Posle OTK-1/HLD-1 | Posle D1-UPIS |
|---|---|---|---|
| Snimanje otkupa, prosek | 12,5 s | 10,5 s | **≈5,6 s** (prvo 6,1 s; ostala 4,6–7,1 s) |
| Pune provere knjige po snimanju | — | 15 (4,8 s) | **0**; 49 nastavaka u 4 snimanja, ≈32 ms svaki |
| Upis (`ScrSave` van sekcija), po snimanju | — | ≈3,7 s | ≈1,2 s |
| Hladnjački lanac, po hladnjačkom snimanju | 6,5 s | 4,5 s | 2,3 s |
| PDF otkupnog lista | 2,27 s | 1,27 s | 0,94 s |
| Osvežavanje posle snimanja | 2,8 s | 2,8 s | 2,2 s |

- Jedina puna provera u sesiji je bila prvo čitanje posle otvaranja sveske
  (329 ms, klik pre prvog snimanja). Posle nje knjiga se samo nastavlja.
- Unosi nisu identični onima iz §3a, a snimanja je četiri. Odnos je zato
  pouzdan, decimale nisu.
- **Šta sada nosi snimanje** (≈5,5 s):
  - osvežavanje posle snimanja 2,2 s; u njemu KPI dvaput računa
    `ReportSaldoOM` (UI-1) i ide kaskada `Change` na OM-u (UI-7, 12× ≈170 ms
    u 4 snimanja);
  - ostatak hladnjačkog lanca ≈1,9 s po hladnjačkom snimanju, neraščlanjeno;
  - upis ≈1,2 s; PDF 0,94 s;
  - nastavak dokaza ≈0,4 s i šema-kapija knjige ≈0,16 s po snimanju.

**Raščlanjavanje osvežavanja posle snimanja** (merne tačke `REF.*`, `NAV|*`,
`OPO.*`, `KPI.*`, `GRID.*`, `LNC.*`; 20.11., 4 snimanja, po snimanju):

| Deo | Trajanje | Šta je unutra |
|---|---|---|
| Osvežavanje ukupno | 2,25 s | |
| Značke menija | 1,05 s | **Bankovni nalozi 0,73 s** (`BuildBlokIsplataList` za svaki prikaz, da bi se prebrojali blokovi); Oporavak 0,32 s (lanci 0,10, palete 0,08, blokovi 0,08, prijemnice 0,05) |
| Liste konteksta | 0,55 s | skoro sve je kaskada `Change` na OM-u, 3× (UI-7): handler, KPI ponovo |
| KPI pločice | 0,36 s | saldo OM-a (`ReportSaldoOM`); „danas” se računa 4× po snimanju |
| Prikaz i mreža | 0,23 s | učitavanje 0,18, crtanje 0,04 |

- **Hladnjački lanac**, po hladnjačkom snimanju: otpremnica 1,22 s,
  prijemnica 0,82 s, zbirna 0,21 s, ostalo malo.
- Rez za ovo su UI-7 i UI-1 iz §3.

**Izmereno posle UI-7 i UI-1** (20.11., 3 snimanja, 2 hladnjačka i 1 obično;
CSV, po snimanju):

| | Pre | Posle |
|---|---|---|
| **Osvežavanje ukupno** | 2,25 s | **1,46 s** (−35%) |
| Liste konteksta (kaskada OM-a) | 0,55 s | 0,01 s |
| KPI saldo / „danas”, broj računanja | 4 / 4 | 1 / 1 |
| Značka „Bankovni nalozi” | 0,73 s | 0,59 s |
| Značka Oporavak | 0,32 s | 0,27 s |
| Mreža, učitavanje | 0,18 s | 0,13 s |
| Čitanja `tblKooperanti` sa lista | ≈193 | ≈16 (537 iz keša u 3 snimanja) |

- Snimanje po vrsti: hladnjačko 7,0–7,1 s → 6,2–6,5 s, obično 4,8–5,0 s → 4,1 s.
  Pre svih rezova (`main`) prosek je bio 12,5 s.
- META je čista u obe sesije.
- „Bankovni nalozi” ostaje najveća značka. Prozor je skinuo samo čitanje
  lista; ostatak je računanje spiska blokova da bi se prebrojali.
- **Šta sada nosi snimanje** (po snimanju):
  - hladnjački lanac: otpremnica 1,2 s, prijemnica 0,8 s, zbirna 0,2 s;
  - upis ≈1,05 s, PDF ≈0,95 s;
  - provera knjige (nastavak) ≈0,55 s i šema-kapija knjige ≈0,21 s
    (13 provera po snimanju);
  - značke 0,86 s, KPI 0,36 s, mreža 0,18 s.

**Raščlanjavanje hladnjačkog lanca** (20.11., 2 hladnjačka snimanja, sekcije
`OTP.*`, `IZD.*`, `PRJ.*`, `LNCP.*`; po hladnjačkom snimanju):

| Korak | Ukupno | Najveći delovi |
|---|---|---|
| **Lanac ukupno** | 2,38 s | |
| Otpremnica | 1,22 s | izdavanje 0,70 s (ponovna provera članova 0,29 s, knjiženje ambalaže 0,40 s); upis izvora 0,15 s; očekivanje 0,13 s; commit 0,11 s |
| Prijemnica | 0,83 s | upis dokumenta 0,48 s (od toga knjiga ≈0,19 s); šema 0,07 s; priprema (izvori, zbir po klasi, cene) 0,16 s; palete 0,06 s |
| Zbirna | 0,23 s | neraščlanjeno |

- **Nema jednog uskog grla.** Najveća zajednička stavka je čitanje knjige
  ambalaže: ≈19 čitanja po hladnjačkom snimanju. Svako ponovo uzima tabelu sa
  lista, traži 12 kolona i radi nastavak dokaza (0,62 s), a 15 puta i
  šema-kapiju (0,24 s, `SchemaReadyOrFail` čita zaglavlja i formate kolona).
- **Kandidati za sledeći krug** (nisu rađeni):
  1. prozor čitanja oko kapija u `PrenesiAmbalazu` (idempotencija, manjak,
     prazne gajbe, obaveza), pre prvog upisa. Kapije tada dele jedno čitanje i
     jedan dokaz. Procena 0,4–0,6 s po hladnjačkom i 0,15–0,3 s po običnom
     snimanju. Dira jezgro pisca knjige: treba mu sopstveni EH koji zatvara
     prozor na svakom izlazu, pa ide kroz `pre-flight`;
  2. šema-kapija knjige jednom po potezu umesto po čitaocu. Kapija čuva
     pozicioni upis, pa i to ide kroz `pre-flight`;
  3. ponovna provera članova u `OtpIzdaj` odmah posle iste provere u
     `CreateOtpremnicaIzIzvora_TX` (0,29 s): odluka D2 je tvoja;
  4. značka „Bankovni nalozi” (0,59 s po svakom snimanju).
- **Snimanje D1-IZV ne dira.** U putu upisa nema prozora čitanja, a kapija
  deficita dokazuje knjigu pre svake odluke. Isti dokaz u upisu (memo po
  transakciji, D1) je sledeći najveći pojedinačni dobitak, oko 4,5 s po
  snimanju. Ide kroz `pre-flight` posebno, jer dira kapiju upisa.

**Prozor Excela u sesiji „posle".** Posle svakog snimanja je list Excela
izlazio ispred forme, a ikonica na traci vraćala je formu.
- Redosled iz CSV-a: „Otvori Excel" je bila 2. radnja sesije, pre svih deset
  snimanja, pa je Excel tokom svih snimanja bio vidljiv.
- Bez „Otvori Excel" pojave nema (operater, 20.11.).
- Diff ove grane prema „pre" je pročitan ceo. Ne dira vidljivost aplikacije,
  aktiviranje prozora ili lista, `ScreenUpdating`, prikaz forme ni PDF.
- Zato ovo nije regresija ove grane. Pojava je zapisana u §6.

## 4. Odluke koje su tvoje

Kritičar je našao mesta gde se nalazi iz različitih oblasti ne slažu oko leka.
Pre izmene treba izabrati jedan mehanizam (pun tekst je na kraju
`02-katalog-nalaza.md`):

- **D1 — knjiga ambalaže: memo ili prozor keša.**
  - Memo u `modAmbalaza` (OTK-2, AMB-INT-4) važi dok traje `PrenesiAmbalazu` /
    blok merenja. Skida i čitanje, i `KnjigaIntegritet`, i šema-kapiju. Ušteda
    je oko 60–80%.
  - Prozor keša (AMB-1, ZAJ-D1, AMB-ISP-1, AMB-STO-2, AMB-ODL-3) skida samo COM
    čitanje, oko 5–15% po AMB-STO-2.
  - Memo je jači, ali dira kritičnu invarijantu (AMB-INV-04/-07/-09/-10/-16 i
    sabotaže `amb-pisac-*`).
  - **Predlog:** prvo izmeriti udeo `KnjigaIntegritet` naspram čitanja. Ako
    integritet dominira, ide memo, kroz `pre-flight`.
- **D2 — mapa članstva otpremnice.**
  - OTP-3 predlaže da se mapa izgradi jednom po potezu.
  - IZM-D1 to izričito isključuje u `DodajOtpremnicaIzvor_TX`,
    `CreateOtpremnicaIzIzvora_TX`, `OtpIspravi` i `LncSastavi`, jer mapa posle
    `AppendRow` više nije tačna.
  - **Predlog:** IZM-D1, jer je najuži i jedini eksplicitno čuva tačnost: mapa
    jednom samo u read-only revalidaciji (`OtpIzmeniDraft`, `OtpIzdaj`).
- **D3 — memo izvedenih struktura** (`StavkeOtkupaRedovi`,
  `ZbirStavkiPoOtkupu`).
  - UI-2 bi keširao i `ZbirStavkiPoOtkupu`, koji vraća `Dictionary`.
  - IZV-3 to zabranjuje dok se ne pregledaju svi čitaoci, jer je deljeni
    `Dictionary` promenljiv.
  - **Predlog:** IZV-3, samo funkcije koje vraćaju niz, uz brisanje na svaki
    upis.
- **D4 — AutoSave i warm.** Oba su na +60 s. Najmanja izmena je razmaknuti ih
  (npr. warm na +90 s) ili pustiti warm tek posle AutoSave-a. Ovo menja trenutak
  pozadinskog posla, ne podatke. Grana „0 s" (neprekidan unos duži od 10 min)
  ostaje.
- **D5 — prozor keša unutar transakcije** (ZAJ-01). Ugovor je da je prozor
  read-only. Svi predlozi ovde zato drže prozor oko read-only blokova: kapije
  pre TX, merenje posle upisa, glava lanca. Prozor preko `AppendRow` bio bi
  promena ugovora `modDataAccess` i ide samo kroz `pre-flight`.

## 5. Tajmer (detalji: `03-tajmer.md`)

**Urađen 19.11.2026** na grani `claude/merenje-modperf` (testovi 392–398, od
20.11. 396–402), uz
granu `claude/merenje-modperf-pre` za merenje „pre". Uputstvo: `04-merenje.md`.
Ispod je prvobitni dizajn.

- **Nov modul `modPerf`, jedan `.bas`.** Ekvivalent ne postoji: postoje pet
  lokalnih `Timer` razlika rezolucije oko 15,6 ms, bez ugnježđavanja, bez
  brojača i svaka sa svojim izlazom. Od njih se preuzima:
  - prag iz `modStornoImpact`;
  - prelaz ponoći i `Str$` iz BFP-a;
  - `#If VBA7 ... PtrSafe` iz `modDataAccess`;
  - folder `Log` i `LogInfo` iz `modLogError`.
- **Sat:** `QueryPerformanceCounter` kroz `Currency`. Isto radi u 32- i
  64-bitnom Office-u. Kad QPC ne radi, rezerva je `Timer`.
- **Prekidač po mašini:** `tblLocalConfig PERF_MERENJE = NE | DA | DETALJ`.
  Čita se jednom na startu. Može i za jednu sesiju kroz Alt+F8 `Perf_Ukljuci`.
  Kad je isključen, svaki hook je jedno čitanje `Boolean`-a
  (`If gPerfOn Then ...`).
- **Izlaz:**
  - `<folder sveske>/Log/Perf_YYYY-MM-DD.csv`, separator `;`, decimalna tačka.
  - Ništa ne ide u tabele sveske, pa nema novog pisca (A11).
  - Nema mreže.
  - U test režimu se ništa ne piše.
- **Šta meri:**
  - vreme po akciji operatera, i ukupno i aktivno (bez čekanja na MsgBox);
  - sekcije sa self i ukupnim vremenom;
  - brojače: `GetTableData` po tabeli sa brojem ćelija, `GetColumnIndex`,
    snapshot po tabeli, `KnjigaZaCitanje`;
  - META proveru `TableCacheDepth = 0`.
- **70 tačaka merenja** su vezane za ID nalaza.
- **Protokol:**
  - samo kopija prave sveske, nikad produkciona;
  - 13 scenarija (S1–S13), 1 hladan i 10 toplih prolaza, A/A provera šuma;
  - kriterijum: nalaz je potvrđen kad nosi ≥10% aktivnog vremena ili ≥100 ms;
  - procena trajanja za operatera: 2–4 h.
- **Provereno ovde:** skica ubačena u kopiju repoa → `vba_check` je čist nad
  celim `src-vba/`, a `who_writes --check-ownership` vraća exit 0.
- **Nije provereno:**
  - Compile u Excelu;
  - ijedan test (predloženo šest `T_Perf_*`);
  - QPC na tvojoj mašini;
  - hook-ovi u postojećim modulima. Oni su samo opisani, nisu ubačeni.
- **Najveći operativni rizik:** u kopiji sveske cloud sync i monitoring moraju
  biti ugašeni, inače kopija može da gura podatke u produkciju.

## 6. Usput uočeno, nije brzina i nije verifikovano

Navedeno da se ne izgubi. Ništa od ovoga nije provereno testom.

- **Žurnal posle rollback-a (ZAJ-09):** `WriteJournalRow` se piše unutar TX, a
  rollback ga ne vraća. Žurnal zato može da sadrži red koji u svesci ne
  postoji.
- **Upis mimo `modDataAccess` (UI-D2):** `BulkPushPendingForStanica`
  (`modStanicaLock.bas:426`, čitanje i upis oko `480`) piše u `tblOtkup` direktno, pa ne prolazi
  invalidaciju keša.
  - Uz to, `who_writes` broji samo `AppendRow` / `UpdateCell` / `DeleteRow` po
    imenu, pa ovaj pisac verovatno nije vidljiv A11 proveri. Nije provereno.
- **Zastarela značka uvoza izvoda:** `modScrBankaUvoz.Kpi` nema proveru
  `DataGeneracija`, pa značka može ostati stara posle upisa sa drugog ekrana.
- **Manjak u prijemnici:** u `PrijemnicaUpisi` se manjak za ponovni pokušaj
  računa nad stanjem PRE transakcije (`modDokUnos.bas:1192`).
- **Heartbeat zaključavanja stanice:** na svakih 90 s radi sinhroni GET i PUT ka
  Google listu u UI niti, sa timeout-om 10/30 s (`modStanicaLock.bas:36`,
  `256-283`). Klik koji padne u taj trenutak čeka. Nije na putu tri toka, ali
  vredi jedno merenje.
- **Zastareo opis u `.claude/rules/otkup-i-dokumenta.md`:** pominje
  `SaveOtkupMulti_TX` (obrisan, `modOtkup.bas:1949`; pisac je `CreateOtkup_TX`)
  i „`AutoChainHladnjaca` pod `On Error Resume Next`". Lanac se zove
  `AutoLanacHladnjaca` i ima sopstveni EH. Ispravka ide **isključivo kroz
  zaseban process PR**.
- **Posle „Otvori Excel" Excel ostaje vidljiv i prekriva formu** pri svakom
  snimanju (merenje 20.11., §3a).
  - Excel se ponovo sakriva samo povratkom preko kartice „Nazad u aplikaciju"
    (`modUiFaze.NazadUAplikaciju`).
  - Verovatno ga ispred forme dovodi PDF otkupnog lista, koji pri snimanju
    preuzme fokus. Mehanizam nije potvrđen.
  - Da li se isto dešava na `main`-u nije provereno. Kod ove grane ne dira
    prozore.

## 7. Šta nije provereno

- **Izmereno je samo ono što je u §3a** (20.11., jedna kopija prave sveske,
  nivo 1, malo ponavljanja). Sva ostala vremena u ovom dokumentu su i dalje
  procene.
- **Nepoznato je:**
  - veličina tabela;
  - da li sveska ima volatilne formule (utiče na OTK-16 i GRP-7: upis u šablon
    sa `Calculation` na automatskom);
  - da li `FormatNeslaganje` na pravoj svesci svaki put okida
    `PrimeniFormateKanona`;
  - koliko traje sam `ThisWorkbook.Save`. AutoSave akcija traje 7,3–7,5 s
    (§3a), ali koliko od toga je `Save` nije razdvojeno;
  - da li je sveska na lokalnom disku, mreži ili OneDrive-u.
- **UI-7:** da `Change` na OM-u pri snimanju zaista okida potvrđeno je merenjem:
  3 puta po snimanju, ≈263 ms svaki (§3a). Koji tačno korak ga diže (`.Clear`,
  programsko punjenje ili izbor) nije razdvojeno.
- **Delovi koje nijedna analiza nije pročitala do kraja** (spisak po oblasti je
  u katalogu, odeljci „Šta analiza ove oblasti NIJE pokrila"). Glavni:
  - `CreateZbirna_TX` / `ZbirnaUpis`;
  - `IspravkaOtpremnice_TX`;
  - revers i novac (`ReversValidiraj`, `UpisiReversAmbalaze_TX`, ...);
  - `LncRasturi`, `LncPreveziPrijemnicu`;
  - liste Dokumenata osim `RedoviZaTip`;
  - `ReportKarticaRobaRekap`, `ReportKarticaAmbalaze`;
  - `AutoFillCena` / `GetVazecaCena` u `ClearForm`.
- **Dodatni nalazi verifikatora (14) i nalazi iz rupa (22)** nisu nezavisno
  verifikovani. Od njih sam u kodu proverio samo HLD-1.

## 8. Pitanja za tebe

**Odgovoreno 19.11.2026** (pitanja 1, 2, 3 i 5; preračun je u §1a):
- otkup do oko 10.000 redova, stavke do 15.000, ambalaža do 50.000;
- otpremnica i zbirna znatno manje od otkupa;
- bez malina režima pojedinačni otkupni list je obavezan, a u malina režimu
  grupni jeste, a pojedinačni nije;
- sve firme koriste hladnjačku stanicu, neke malina režim;
- Office je najčešće 64-bitni.

**Otvoreno:** pitanje 4 (broj kooperanata u bloku isplata) i tačan broj
otpremnica i zbirnih. §1a pretpostavlja 1.000–3.000.

Pitanja kako su postavljena:

1. Koliko redova otprilike imaju `tblOtkup`, `tblOtkupStavke`, `tblAmbalaza`,
   `tblOtpremnica`, `tblZbirna`, `tblKooperanti` i `tblNovac` u sezoni?
   - #1 i #9 su kvadratni: pri 1.000 redova su neprimetni, a pri 10.000 su
     desetine sekundi.
   - Brojeve daje i tajmer (VEL redovi), ali i gruba procena iz glave pomaže da
     se izabere prvi rez.
2. Da li se otkupni i grupni list moraju štampati ili praviti kao PDF uz
   **svako** snimanje? Ako ne, korak 1 iz §3 je besplatna dobit.
3. Da li radite sa hladnjačkim stanicama (RUTA_LANAC) i u malina režimu? Ako ne,
   #9 i #13 padaju na dno.
4. Koliko je otvorenih kooperanata u bloku isplata (značka BANKA_NALOZI)?
   ZAJ-D2 se kreće od 0,3 do 10 s zavisno od toga.
5. Office 32- ili 64-bitni? Važi za test `T_Perf_QpcMonoton`.

## 9. Kako je rađeno

- **Workflow sa 17 agenata:**
  - 6 analiza: snimanje otkupa, snimanje ostalih dokumenata, izmena i
    ispravka, Izveštaji, zajednički sloj, osvežavanje UI;
  - 6 suprotstavljenih verifikacija;
  - dizajn tajmera;
  - kritičar pokrivenosti;
  - 3 istraživanja rupa: hladnjački lanac, grupni otkupni list u malina režimu,
    storno ambalaže i storno u lancu.
- **Pravila za agente:** ciljano čitanje (grep i opsezi linija), svaka tvrdnja
  sa `fajl:linija`, procene uvek označene kao PROCENA.
- **Spojeni duplikati** su navedeni u kolonama tabele iz §2.
- **Preimenovani ID-jevi** iz izlaza agenata su na vrhu kataloga (npr. `DOD-1`
  iz izmene → `IZM-D1`).
