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

## 2. Rang po očekivanoj dobiti

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

1. **Bez ijedne linije koda, odmah** (#5):
   - **Podešavanja → „Otvori PDF posle izrade"** (`PDF_OTVORI_POSLE`) = NE
     (`modPodesavanja.bas:125`). PDF se i dalje pravi i arhivira, samo se ne
     otvara i ne uzima fokus.
   - Gde papir nije obavezan, mod štampe otkupnog ili grupnog lista
     (`OTKUP_PRINT_MODE`, `GRUPNI_OTKUP_PRINT_MODE`; vrednosti
     `PDF;PRINT;PREVIEW;OFF`, `modPodesavanja.bas:104-105`) = OFF. Dok se #1 ne
     ispravi, OFF skida i kvadratnu petlju otkupnog lista.
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
   - #9 HLD-1;
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

## 7. Šta nije provereno

- **Ništa nije pokrenuto u Excelu.** Svako vreme je procena; tačni su samo
  brojevi poziva izvedeni iz koda, i to tamo gde su prebrojani.
- **Nepoznato je:**
  - veličina tabela;
  - da li sveska ima volatilne formule (utiče na OTK-16 i GRP-7: upis u šablon
    sa `Calculation` na automatskom);
  - da li `FormatNeslaganje` na pravoj svesci svaki put okida
    `PrimeniFormateKanona`;
  - koliko traje `ThisWorkbook.Save`;
  - da li je sveska na lokalnom disku, mreži ili OneDrive-u.
- **UI-7:** da li MSForms `ComboBox` diže `Change` na `.Clear` i pri
  programskom punjenju. Bez Excela se ne može utvrditi.
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

Odgovori menjaju rang iz §2:

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
