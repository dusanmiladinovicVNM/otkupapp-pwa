# Refaktor playbook — sprovođenje plana ispravki (jedan paket = jedna sesija)

**Izvori:** `docs/AUDIT_FM_TRIJAZA.md` (detalji svake stavke, dokazi fajl:linija) ·
`docs/KNOWN_ISSUES.md` §8 (AUD registar) · `docs/ROADMAP.md` §10 (Wave plan).
**Model rada:** svaki paket (RF-xx) se radi u NOVOM chatu, na svojoj grani, serijski
(paket se merge-uje u `main` pre početka sledećeg). Jedan paket ≈ jedna sesija.

---

## 1. Prompt šablon za novu sesiju

Zalepiti u nov chat (Claude Code na ovom repou — CLAUDE.md se učitava automatski):

```
Radi paket RF-XX iz docs/REFAKTOR_PLAYBOOK.md.
Pre izmena: pročitaj sekciju paketa u playbook-u, navedene AUD/FM stavke u
docs/AUDIT_FM_TRIJAZA.md i SVE module iz obima (cele, ne isečke).
Radi ISKLJUČIVO obim paketa — ništa van njega. Na kraju: statičke provere,
commit+push na granu claude/rf-XX-<slug>, git komande za preuzimanje,
numerisana Excel test checklista, i ažuriraj Status tabelu u playbook-u.
```

## 2. Radna pravila (važe za svaki paket)

1. **Grana:** `claude/rf-XX-<slug>` sa svežeg `main`-a. Serijski — ne počinjati novi
   paket dok prethodni nije merge-ovan (paketi dele fajlove).
2. **Obim je zakon.** Radi se samo ono što paket navodi. Ako se usput otkrije nov
   problem: NE popravljati — upisati kao novu stavku u `KNOWN_ISSUES.md` §8 (AUD-0xx)
   i nastaviti.
3. **Minimal delta** (CLAUDE.md doktrina): bez novih apstrakcija, bez preimenovanja,
   bez „usputnog sređivanja". Postojeći obrasci su šablon: `RequireColumns`/
   `RequireUpdateCell`, `_TX` + ne-TX core, `BuildPrijemnicaRowData` (upis po imenu),
   runtime paneli, `Poruka("KLJUC")` katalog.
4. **VBA ograničenja:** izvori ostaju 100% ASCII (dijakritika SAMO kroz
   `modPoruke.UpsertPoruke` + `ChrW`); `.frx` se ne dira; nove kontrole samo runtime
   (`Controls.Add` + WithEvents).
5. **Statičke provere pre commit-a:** `file` = „ASCII text" za svaki izmenjen VBA fajl;
   grep ne-ASCII = prazno; balans `Sub/Function/End`; nema duplih `Public` definicija
   (`grep -h "^Public " src-vba/*.bas | sort | uniq -d`); svaki novi `Poruka("X")` ima
   par u `UpsertPoruke`.
6. **Kraj svakog paketa:** commit sa jasnom porukom → push → git komande za
   preuzimanje grane → **numerisana Excel test checklista** (klik-po-klik, fokus samo
   na izmene paketa + navedeni regression suite) → ažurirati Status tabelu (§4).
7. **Merge u main:** tek posle korisnikovog Excel testa (`ImportAllVBA` → Compile →
   checklista). Posle merge-a: `tools/release.sh` po `docs/RELEASE_PROCEDURE.md` kada
   se skupi smisleni skup paketa (npr. posle RF-05, RF-08, RF-14).

## 3. Paketi (redosled izvršavanja)

> Detalji svake stavke (tačne linije, predlog, napor) su u `docs/AUDIT_FM_TRIJAZA.md`
> pod navedenim FM/AUD referencama — ovde je samo obim i definicija gotovog.

### RF-01 — Brisanje balasta [Wave 0 · S]
**Fajlovi:** `src-vba/` (brisanje), `modE2EReleaseGate.bas` (bez izmene — verifikacija).
**Obim:** obrisati `modNovacTest.bas`, `modFakturaTest.bas`, `modLicenceTests.bas`
(ostaju `*Tests` verzije), `modBankaImportParserClipboard.bas`; u `modBankaImport.bas`
obrisati `GetFileNameFromPath2` (poziv :1163 preusmeriti na original); u
`modArrayUtils.bas` obrisati mrtve `GroupBySum`/`SumColumn`; u `modIzvestaj.bas`
mrtav enum `IzvestajTip`; u `frmIzvestaj.frm` mrtve `UpdateUnosButtonState` i
`PrijemniceZaOtpremnicu`. [AUD-016; FM-0027 #6/#7/#34; FM-0029 #28/#29]
**Gotovo kad:** grep pokazuje 0 referenci na obrisano; E2E gate `Application.Run`
imena sada jednoznačna. **Checklista mora reći:** obrisane module ukloniti i ručno u
VBE jednom (ImportAllVBA ne briše komponente).

### RF-02 — modNovac finansijski guardovi [Wave 1 · S/M]
**Fajlovi:** `modNovac.bas`, `modCenovnik.bas`.
**Obim:** (1) `RequireColumns` za svih 17 kolona na vrhu `SaveNovac` (P0, AUD-003);
(2) `ApplyAvansToOtkup/Faktura`: target-active (storno) + target-owner provera +
`_TX` vraća primenjeni iznos umesto no-op `True` (AUD-010; FM-0019 #4/#5/#6/#11);
(3) `AddCena` presence guard (AUD-003). **Regression:** `RunNovacSmokeSuite`,
`RunBusinessFlowProSuite`.

### RF-03 — Storno „lažni uspeh" lanac [Wave 1 · M]
**Fajlovi:** `modStornoFlow.bas`, `modDokumentInvariant.bas`, `modStorno.bas`.
**Obim:** context-guard u 5 nezaštićenih grana; provera svakog paletnog detach
rezultata; provera zbirna relink count-a; invariant existence (`0=0` ne sme proći za
nepostojeću zbirnu); sum-scan error flag; `LookupActiveID` multi-match detekcija;
`StornoNovac` → čitaj `OtkupID` + `UpdateOtkupStatus` (+ snapshot `tblOtkup`)
(AUD-020, AUD-021; FM-0013 #1/#2/#3/#7, FM-0015 #1/#3, FM-0012 #12, FM-0019 #16).
**Regression:** `RunStornoTestSuite` (modTestStorno), `RunBusinessFlowProSuite`.

### RF-04 — modAutoHladnjaca lanac [Wave 1 · S/M]
**Fajlovi:** `modAutoHladnjaca.bas`.
**Obim:** `outBrPrij` postaviti tek posle uspešnog `SavePrijemnica_TX`; backfill
brojeve seedovati postojećim prijemnicama zbirne; proveriti `otpID` i `SaveZbirna_TX`
rezultat (prekid klase + warning); propagirati link grešku u warning lanca; `have`
dopuna posle uspeha (AUD-005; FM-0010 #1/#2/#4/#5/#6/#9). **Regression:** ručni
malina tok + `RunBusinessFlowProSuite`.

### RF-05 — frmDokumenta unos + storno set [Wave 1/2 · M]
**Fajlovi:** `frmDokumenta.frm`, `modDokumenta.bas`.
**Obim:** novac storno broj→`NovacID` rezolucija (kao za fakturu na :3222);
`FillOpenFakture` storno filter; Kl.II checkbox hard blokada kad izvor ima Kl.II;
obavezan smer ambalaže uz `kolAmb>0`; malina auto-zbirna vidljiv pad; prefill
poslednja generacija; `SumByBroj` storno filter; `SaveZbirna` presence guard
(AUD-008/009/022 + delovi AUD-003; FM-0018 #1-#5/#8/#20, FM-0011 #3).
**Regression:** `RunStornoTestSuite` + ručni unos otpremnica/zbirna/prijemnica.

**Urađeno (grana `claude/rf-05-frmdokumenta-fixes-63yqjp`):** svih 6 planiranih fiksa
(novac storno broj→`NovacID` je već rešen u RF-03, pa preskočen) + obavezan smer ambalaže
(UI + core guard u `SaveOMUlaz_TX`). Tokom review-a dodato: `FillOpenFakture` prebačen na
centralni `modNovac.GetOpenFakture`; `SaveZbirna` na column-mapped `BuildZbirnaRowData`;
prefill generacije rešen **eksplicitnom `GeneracijaID` kolonom** (schema — `EnsureSledljivostSchema`)
umesto heuristike po datumu/ID-u, sa anchor-om na `OldDocID` iz correction context-a;
`RequireJedanVlasnikPoBroju` guard protiv storna po nejedinstvenom broju (AUD-052).
Test seam: `ZbirnaIzvorImaKlasuII`, `PickPrefillRows`, `SaveOMUlaz_TX` → `modDokumenta`.
**Ostaje za RF-06+:** pun identitetski storno (`OldDocID → GeneracijaID → redovi generacije`
kroz `Scan*`/`Run*Correction`) — vidi AUD-052 u `KNOWN_ISSUES.md`.

### RF-06 — modIzvestaj ispravnost brojki [Wave 2 · M/L]
**Fajlovi:** `modIzvestaj.bas` (+ `modNovac.bas` per-vrsta cache).
**Obim:** isplate kooperanta vezati za stanicu (`COL_NOV_OM_ID`); kartice red
„Početno stanje"; „nema prijema" oznaka umesto 0%/100% nekonzistentnosti; kupac
per-vrsta raspodela uplate srazmerno stavkama; eksplicitni `Select Case` u Report*
dispatch-u (Else → Empty/greška) (AUD-023; FM-0028 #1/#3/#5/#6/#13/#14 + P2 stavke
#2/#9/#10/#11 po proceni sesije). **Regression:** `RunIzvestajTests` + uporedni
pregled izveštaja pre/posle na istim podacima (checklista mora dati očekivane razlike).

**Urađeno (grana `claude/rf-06-izvestaj-brojke-wquclc`):** svih 5 planiranih fiksa
+ P2 #9 (dolazi „besplatno" uz #3 — atribucija ide po `COL_NOV_OM_ID` reda) i #10
(neraspoređena agrohemija van UKUPNO stanice). Uvedeni deljeni računski seam-ovi u
`modIzvestaj`: `NovacRedPripadaStanici`, `ManjakStavka` (deljen između `ReportOtkupRobaOM`
i `ReportManjak` — „rule of two", ista brojka na dva mesta), `PrijemZaZbirnu`,
`KarticaRezultatSaPocetnim`, `KarticaAmbRezultatSaPocetnim`; u `modNovac`
`BuildVrstaFakturaCache` → `BuildFakturaVrstaUdeoCache` + čiste `RaspodeliPoUdelima`
i `ZaokruziNovac`. Nov `RunIzvestajTests` (tvrd gate — `Err.Raise` na pali assert).

**Dopuna posle review-a (isti paket, 2. commit):**
- **`BrojZbirne` nije identitet ni u report sloju** (posledica AUD-052 koju je RF-05
  već dokazao na storno putanji). `modHelpers.BuildManjakDict` je spajao zbirne i
  prijemnice **isključivo po poslovnom broju**, a `ReportManjak` je istu grešku
  imao i u sopstvenoj, dupliranoj agregaciji → dve aktivne zbirne istog broja su
  sabirale tuđu prijemnu količinu (pogrešan prijem, manjak, %). Sada je
  `BuildManjakDict` scoped na vlasnika (`ZbirnaVlasnikKljuc` = `broj|vozac|kupac` —
  **ista definicija vlasnika koju koristi `modStorno.RequireJedanVlasnikPoBroju`**,
  bez druge paralelne definicije; klasa se dodaje u sledećoj stavci),
  `ReportManjak` više ne duplira agregaciju, a
  `ReportOtkupRobaOM` razrešava vlasnika preko vozača otpremnice (otpremnica nema
  `KupacID`). Nedokaziv vlasnik → **fail-closed** oznaka `nejasan vlasnik`, van UKUPNO.
  Broj sa jednim vlasnikom koristi agregat po broju → nema regresije na starijim
  prijemnicama bez popunjenog vlasnika.
- **`Klasa` mora biti u ključu manjka.** Owner-ključ `broj|vozač|kupac` je i dalje
  spajao **Klasu I i Klasu II istog dokumenta** — auto-lanac hladnjače ih vodi kroz
  ceo lanac odvojeno (zasebna otpremnica/zbirna/prijemnica) ali sa istim brojem,
  vozačem i kupcem. Posledica u malina modu: zbirni prijem obe klase (npr. 900+150)
  dodeljivao se **svakoj** otpremnici → UKUPNO prijem 2× stvarni. Ključ je sada
  `broj|vozač|kupac|Klasa` (`ZbirnaStavkaKljuc`) — ista granularnost koju
  `modAutoHladnjaca.KeyZbrKlasa` već koristi, pa nije uveden nov pojam.
  `#V|` i dalje broji **vlasnike** bez klase (dve klase ≠ dva vlasnika).
  `ReportManjak` zadržava jedan red po dokumentu ali prijem sabira po klasama;
  `ReportOtkupRobaOM` radi srazmeru unutar klase.
- **Finansijsko zaokruživanje raspodele (dva kruga).** `RaspodeliPoUdelima` je prvo
  zaokruživala samo ukupan zbir, pa je 100/3 davalo interne delove 33,3333 → prikaz
  `33,33 × 3 = 99,99` uz UKUPNO `100,00`. Prva popravka (poslednji ključ nosi ostatak
  posle zaokruživanja) rešila je zbir ali uvela **negativan cent**: kad prethodni delovi
  zaokruživanjem pređu cilj, poslednji ode u minus (`0,03` na 5 jednakih vrsta → `−0,01`).
  Konačno rešenje je **largest-remainder u celim parama** — `Int` idealnog udela + višak
  para po najvećim ostacima — jer jedino ono drži **obe** invarijante odjednom
  (zbir == iznos **i** nijedan deo < 0; clamp na nulu bi razbio prvu). Pare u `Double`,
  ne `Long` (Overflow preko ~21,4 mil.). Vidljiva promena: kod jednakih udela višak pare
  dobija **prvi** ključ umesto poslednjeg.
- **Test gate:** `RunIzvestajTests` sada podiže grešku na pali assert (konvencija iz
  RF-14) i dobio je **tri end-to-end testa nad tabelama** (dve zbirne istog broja,
  dva kupca, 900 vs 1500 kg — iz ugla `ReportManjak` i `ReportOtkupRobaOM`; plus Klasa I+II
  istog dokumenta 1000/900 i 200/150), jer
  seam testovi po definiciji ne mogu da uhvate grešku u samom table-join-u. E2E rade
  u `clsTransaction` sa snapshot-om i uvek se rollback-uju. Novčane provere idu na
  nivou centa (`IzvChkEqC`) — tolerancija 0,01 bi propustila nezaokruženo 33,3333.
**Svesno NIJE uzeto (ostaje RF-07 / UI paket):** #2 header „Amb. (trenutno stanje)"
i #11 labela „OM AVANS (promet perioda)" — čist UI tekst; #11 dodatno lomi
`modTestStorno` T29 koji taj red traži po literalu, pa ide zajedno sa header izmenama.
Vidljiva poruka za nevalidnu kombinaciju (core sad vraća `Empty`) traži `CleanFail`
fix iz RF-07 — do tada je ishod čista prazna lista, ne pogrešne brojke.
**AUD-013 (`MatchesFilter`):** prebrojani SVI `clsFilterParam.Init` pozivi u `src-vba/` —
operatori su literali iz podržanog skupa, nijedna report putanja ne zavisi od
`Case Else`; grana je nedostižna u produkciji → flagovano u `KNOWN_ISSUES`, scope
se ne širi (fix dira ceo `ExcludeStornirano` sloj).

### RF-07 — frmIzvestaj freshness + revers [Wave 2 · M] — ✅ urađeno
**Fajlovi:** `frmIzvestaj.frm`, `modIzvestaj.bas`, `modKarticaDetalji.bas`, `modPoruke.bas`, `modIzvestajTests.bas`.
**Review (REQUEST CHANGES na `71355a4`) — sve tri stavke prihvaćene i ispravljene:** matrica vs core na `Kupac+Zbirni+Prosečna cena`, invalidacija konteksta pri promeni entiteta bez dostupnog izbora, kanonski ključ tipa ambalaže.
**Review 2 (REQUEST CHANGES na `4f7b600`) — prihvaćeno:** grupni ključ pregleda ambalaže dopunjen `DokumentTip`-om (uz-otkup revers je bio nedostupan za štampu).
**Review 3 (`/code-review` max-effort: 0 correctness bugova, 3 LOW) — sve tri rešene:** `m_periodDirty` reset u `InvalidateReportContext`, uklonjena mrtva grana u status poruci, uklonjena puna kopija ledgera u revers štampi.
**Obim:** status/štampa iz `m_curOd/m_curDo` (+ „nije osveženo" na izmenu datuma);
`CleanFail` čisti listu + vidljiva greška; zbirni tabovi 5/6/7 samo za validne tipove;
`StampajReversAmbDok`: `ExcludeStornirano` + tip ambalaže u match; `KarticaDetalji_Clear`
na promenu taba (AUD-024, AUD-012, deo AUD-027; FM-0029 #1-#5/#14/#15/#16/#19, FM-0030 #1).
**Regression:** ručni prolaz kroz tabove + revers štampa za slučaj sa storniranim redom.

**Naučeno (RF-07):**
- **Re-verifikacija pre koda (§2.7):** svih pet nalaza je re-lociran po IMENU posle RF-06
  (linije iz audita v2.24.0 su pomerene) i **svi se i dalje reprodukuju** — RF-06 je dirao
  `frmIzvestaj` samo u `Generate*Report` prikazu oznake `nema prijema`.
- **`MsgBox` pod `Resume` se ne radi.** Prijava greške iz `CleanFail` bloka ide **posle**
  `Resume CleanExit` (u samom `CleanExit`, kad su `EndTableCache` i `ScreenUpdating=True`
  već odrađeni). Poruka se prenosi kroz `m_genFailMsg`/`m_genFailTab`, ne kroz `Err`
  (`Resume` briše `Err`). Isti razlog zašto EH u projektu prvo prepiše `errNum/errDesc/errSrc`.
- **Tab-matrica mora da fiksira i ono što OSTAJE.** Test `T_TabMatrica` ne proverava samo
  da nevalidne kombinacije nestanu, nego i **ceo pojedinačni režim** — prelazak sa hardkodiranih
  `Pages(n).Visible = True` na matricu je tačno ona izmena koja može tiho da skloni ispravan
  tab. Matrica je izvedena iz `Select Case`-ova u `Report*`, ne iz zatečenog UI-ja.
- **Ključ reversa je trojka, ne par.** `ReportAmbalazePojedinacni` red pregleda već grupiše
  po `DokumentID|TipAmbalaze`, pa je jedini konzistentan ključ za rekonstrukciju
  `DokumentID + DokumentTip + TipAmbalaze`. Prazan tip je **legitimna grupa** (red bez tipa),
  ne wildcard — inače bi prazan tip pokupio sve tipove i vratio baš onaj bug koji se zatvara.
- **`_Change` handler na `.frx` kontroli nije `WithEvents`.** Zamka #11 se odnosi isključivo na
  `Private WithEvents` **deklaracije**; obični `txtDatumOd_Change` je isti obrazac kao zatečeni
  `cmbEntitet_Change`/`lstKartica_Click` i ne dira code-merge. (Forme su i inače reinstall-only
  zbog module-level `As MSForms.*` — zamka #20.)
- **Novi `Poruka()` ključevi = obavezan `EnsurePoruke` posle importa.** RF-05/RF-06/RF-28 su svi
  bili „bez novih ključeva"; RF-07 uvodi 7, pa release note to mora eksplicitno da nosi —
  bez `EnsurePoruke` status i poruke prikazuju `[KLJUC]`.
- **Matrica dostupnosti mora da se meri prema ULAZU koji joj pozivalac šalje, ne prema
  postojanju grane.** `Kupac + Zbirni + Prosečna cena` je prošao prvu verziju matrice jer
  `ReportProsecnaCena` *ima* `Case "Kupac"` — ali zbirni režim šalje `entitetID = ""`, a ta
  grana ide kroz `GetPrijemniceByKupac` koji bezuslovno dodaje `KupacID = ""` i vraća prazno.
  `OM` je imun samo zato što grana glasi `Case "OM", ""` i eksplicitno hvata prazan ID.
  **Posledica za testove:** matrični `True/False` assert ne dokazuje da ponuđeni tab može da
  vrati podatke — za svaku „DA" ćeliju koja zavisi od praznog `entitetID`-a treba E2E nad
  tabelama (`T_E2E_ProsecnaCenaZbirniKupac`, seed dva kupca u sentinel prozoru, pada u oba smera).
- **Guard koji preskače posao mora prvo da invalidira ono što ostavlja.** `AutoRefresh` izlazi
  kad novi entitet nema nijedan izbor; bez invalidacije liste prethodnog entiteta ostaju kao
  „svež" rezultat pod novim identitetom — **ista klasa greške koju paket zatvara**, samo što je
  period ispravan a entitet pogrešan. Invalidacija ide na **početak jedinog ulaza** (`AutoRefresh`),
  pre guarda, a ne u svaki `tgl*_Click`.
- **Ako se period pamti, mora i identitet.** `m_curOd/m_curDo` su rešili „stari podaci, nov period";
  isti argument važi za entitet — otud `m_curEntLabel`/`m_curEntName` i `m_curTip` kao izvor za
  naslov štampe i za izbor zaglavlja kolona, umesto živog `GetActiveEntitetTip()`/`cmbEntitet.value`.
  Uz to print guard `AktivanTabGenerisan()`: štampa se samo tab koji je **stvarno** generisan.
- **Matcher nad nečim što je već grupisano mora da deli SIMBOL normalizacije, ne „istu ideju".**
  Pregled je grupisao po sirovom tipu ambalaže, a novi `ReversRedPripada` poredio `Trim +
  vbTextCompare` — „Letvarica"/„letvarica" dalo bi dva reda pregleda a svaki revers zbir oba,
  tj. tiho vraćanje baš onog mešanja koje se zatvara. Rešenje je jedan `AmbTipKljuc` koji zovu
  obe putanje.
- **Izjednačavanje matchera i grupisanja znači CEO ključ, ne samo normalizaciju delova.**
  Prva runda review-a izjednačila je normalizaciju tipa ambalaže (`AmbTipKljuc`) — ali je
  grupni ključ pregleda i dalje bio `DokumentID + TipAmbalaze`, dok matcher traži
  `DokumentID + DokumentTip + TipAmbalaze`. Kolizija je na **normalnoj putanji**:
  `SaveOtkup` namerno piše isti `otkupID` i isti `tipAmb` pod `Otkup` (primljene pune) i
  pod `OM-Izlaz-Koop` (izdate prazne), pa su se spajali u jedan red čiji ref-ključ nosi tip
  prvog zapisa — i revers izdate ambalaže **nije se mogao odštampati iz pregleda**.
  Pouka: kad se dve putanje proglase saglasnim, to mora biti svojstvo koda (isti izraz
  ključa), ne tvrdnja u komentaru; i mora se proveriti šta **pisci** upisuju pod tim ključem,
  ne samo kako ga čitaoci porede.
- **Invarijanta stanja se ne sme oslanjati na redosled provera koje je čitaju.**
  `InvalidateReportContext` nije resetovao `m_periodDirty`; radilo je tačno samo zato što
  grana `Not m_hasGenerated` u `UpdateStatusLabel` izlazi **pre** grane `m_periodDirty`.
  Nije bio živ bug, ali je „radi na sreću" — reset je dodat da invarijanta („nema
  generisanog perioda → nema ni neosveženog u odnosu na šta") važi nezavisno od poretka
  čitalaca. Isto je uklonjena i posledična mrtva grana `If m_hasGenerated Then` u toj poruci.
- **Reuse deljenog helpera nije bezuslovan — merilo je šta helper radi na toj putanji.**
  `ExcludeStornirano` je ispravan izbor u report sloju, ali u one-shot print handleru pravi
  **još jednu kopiju celog `tblAmbalaza`** samo da bi se odštampao jedan dokument, dok petlja
  ionako prolazi ceo ledger. Zato je provera inline (`AmbRedStorniran`) sa **identičnim
  pravilom** (`FilterArray` `"<>"` `"Da"` → `CStr` poređenje, bez trima i case-fold-a; kolona
  koje nema = nema storna) — ista odluka koju `ReportAmbalaza` već nosi i dokumentuje za istu
  tabelu. Kad se pravilo duplira radi performansi, komentar mora da imenuje izvor istine.
- **Modul-level `Const` mora u deklaracionu sekciju — VBA ne kompajlira `Const` između
  procedura.** `IZV_TAB_*` su bile stavljene tik iznad `IzvestajTabDostupan` („uz funkciju
  koja ih koristi"), na sredini `modIzvestaj` — što je prirodno mesto i tačno pogrešno.
  Operater je to morao ručno da premesti u Excelu da bi `Compile` prošao. Nijedna od mojih
  statičkih provera (balans, ASCII, dupli `Public`) ovo ne hvata, jer je sintaksno ispravan
  red na nedozvoljenoj poziciji. **Dodato u `CLAUDE.md` §4 kao pravilo i u §5 kao obaveznu
  statičku proveru;** repo-wide skener potvrdio je da je ovo bio jedini takav slučaj.
**Svesno NIJE uzeto:** FM-0029 #11 (status broji UKUPNO/summary redove), #17 (revers bez
datuma tiho postaje `Date`), #26 (redosled perioda nevalidiran), #9 (prazan `entitetID`) —
svi P2/P3 van zadatog obima paketa.

### RF-08 — modFaktura + faktura štampa [Wave 2 · S/M]
**Fajlovi:** `modFaktura.bas`, `modPrint.bas`.
**Obim:** `CreateFaktura`: poređenje kupca prijemnice + `rows.Count=1` guard +
`CreateFaktura` na Private; `FillFakturaSablon` cleanup `.UnMerge`; blokada reprint-a
storniranog otkupa (ukinuti raw fallback :334) (AUD-011, AUD-027; FM-0034 #1/#2/#3,
FM-0031 #3/#19). **Regression:** `RunFakturaSmokeSuite`; ručno: faktura sa 3 stavke
posle fakture sa 1 stavkom (merge test).
**Naučeno (RF-08):** (1) **„Isti obrazac kao ostali `Fill*`" nije uvek prenosiv.** Ostalih pet
`Fill*` šablona radi `ws.cells.UnMerge` bezbedno **jer pre punjenja ruše i grade ceo list**;
`FakturaSablon` je perzistentan (`EnsureFakturaSablon` gradi ga jednom, po `H1` `LAYOUT_VER`)
i u zaglavlju ima namerne merge-ove (`FakKupac`, seller header, naslov). Blanket `UnMerge` bi
ih trajno pokidao i sledeća faktura bi izašla sa razbijenim zaglavljem. Obrazac je zato preuzet
**opsegom koji se puni** (isti `startCell`..`Offset(80, 5)` koji cleanup već koristi), ne celim
listom. Pouka: pre kopiranja obrasca proveri i **pretpostavku** pod kojom on važi na izvornom
mestu. (2) **Filter koji sakriva podatak sakriva i razlog za blokadu.** Reprint je zvao
`ExcludeStornirano` pa tražio red — storniran red nije nađen, `brDok` je ostao prazan, i raw
fallback (`If ids = "" Then ids = otkupID`) ga je ipak odštampao. „Nije nađen" i „poništen" su
iz filtriranog niza **nerazlučivi**, pa provera mora da ide nad sirovom tabelom i **pre**
filtriranja. (3) **Menjanje potpisa test-helpera je deo scope-a fixa:** `AppendTestPrijemnicaRow`
nije upisivao `KupacID`, pa bi novi ownership guard oborio 4 postojeća testa. Dodat je obavezan
parametar (ne `Optional` sa defaultom — test red bez kupca s pravom više ne sme da prođe).
Provereno je i da svih 5 `CreateFaktura_TX` fixtura u `modBusinessFlowProTests` ide kroz
`SavePrijemnica_TX(… TEST_KUP_ID …)`, pa ta suite ne regresira.
**Naučeno iz code-review-a (RF-08, tri korekcije):** (4) **`Boolean` helper koji nosi sigurnosnu odluku bira smer otkaza svojim `EH`-om.** Prva verzija guarda (`OtkupStorniranZaStampu`) vraćala je `False` na svaku grešku — `Exit Function` posle `LogErr` je uvek „dozvoli". To je fail-**open** baš u schema-corruption scenariju, i to najgorem: `ExcludeStornirano` bez kolone `Stornirano` **takođe** vraća sirovu tabelu, pa bi guard koji „ne zna" i filter koji „ne filtrira" zajedno pustili poništen dokument na papir. Sigurnosne provere se pišu kao `Require*` koje **pucaju**, a pozivalac hvata i prikazuje; `GetColumnIndex` → `RequireColumnIndex`; „prvi pronađeni red" nije dokaz kad ID može biti dupliran. (5) **Fiksna „dovoljno velika" granica u rendereru je invarijanta koju niko ne održava.** `.UnMerge` je bio tačan alat na netačnom opsegu: cleanup je pokrivao 80 redova, a broj stavki nije ograničen ni u UI ni u `CreateFaktura` ni u rendereru — 81 → 82 stavke je i dalje lomilo isti bug (merge preživi, pisanje preko njega → `EH` → `Nothing` → tiho izostala štampa), plus `NumberFormat = "@"` je otpadao od 81. reda. Ili se maksimum enforce-uje na **sva tri** mesta, ili se opseg računa dinamički — pola rešenja pomera bug, ne uklanja ga. Uz to: regresioni scenario mora da gađa **granicu**, ne udoban slučaj (`1 → 3` ne dokazuje ništa o `81 → 82`). **Druga runda istog nalaza:** prvi pokušaj dinamičke granice koristio je `ws.UsedRange` — a `UsedRange` broji i prazne **formatirane** ćelije, dok `EnsureFakturaSablon` formatira ceo list (`ws.cells.Font...`). Granica je time mogla da odleti na poslednji red lista i svaki render bi čistio ~6M ćelija. Zamenjeno sa `SablonLastContentRow` (`Range.Find` `xlFormulas` + `xlPrevious`, samo `A:F`). **Pouka:** `UsedRange` odgovara na „dokle je list dodirnut", a pitanje je bilo „dokle je nešto napisano" — kad granica upravlja **cenom** operacije, semantika njenog izvora je deo ispravnosti, ne detalj. Isto tako: test koji proverava funkcionalni rezultat (`81 → 82` prolazi) ne proverava **koliko ćelija** je obrađeno — za to treba zaseban assert nad granicom. (6) **Hard gate mora da izađe iz aktivnog error handlera.** `RequireFakturaSuiteGreen` pozvan dok je `On Error GoTo EH` još aktivan šalje sopstveni raise u `EH` suite-a: `LogFakturaFatal` doda lažan failure, `FinishFakturaSuite` se izvrši drugi put, operater dobije dupli MsgBox i broj padova veći za jedan. `On Error GoTo 0` pre gate-a — obrazac koji `RunMasterSyncSmokeSuite` već nosi.
**Svesno NIJE uzeto:** AUD-054 (SEF `Err.Raise 0`) — zaseban SEF-hardening; faktura-status /
SEF lifecycle (RF-22); document-snapshot za reprint (TL-006, zaseban roadmap). Fallback za
**aktivan** otkup bez `BrojDokumenta` je namerno zadržan — ukinut je samo put kojim je storno
prolazio.

### RF-09 — Banka import + mapiranje [Wave 2 · M]
**Fajlovi:** `modBankaImport.bas`, `modBankaMapiranje.bas`, `frmBankaImport.frm`.
**Obim:** dedupe ključ + broj računa; 3+ kandidata → `Err.Raise` (umesto subscript
pada); smer guard u Map* funkcijama; Activate auto-map → vidljiv rezultat ili iza
dugmeta; preview/command isti izvor za blok; upozorenje „ručno kupac = avans"
(AUD-014, AUD-025; FM-0022 #1, FM-0023 #8/#14, FM-0024 #1/#2/#3). **Regression:**
`Test_BankParse` po banci + ručni uvoz test izvoda + mapiranje.

### RF-10 — Banka export pregled [Wave 2 · S/M]
**Fajlovi:** `frmBankaExportPregled.frm`, `modBankaExportPregled.bas`.
**Obim:** stale override clamp u `PruneStaleOverrides`; finalna saldo revalidacija u
`GenerisiNalogeCSV`; brojanje stvarno primenjenih avansa (oslanja se na RF-02 iznos)
(AUD-026; FM-0020 #1/#2, FM-0021 #1). **Regression:** ručno: override → izmena
podataka → reload → CSV mora odbiti preplatu.

### RF-11 — frmOtkup / blokovi / kooperant [Wave 2 · M]
**Fajlovi:** `frmOtkup.frm`, `modOtkupBlok.bas`, `modKooperant.bas`.
**Obim:** parcela poređenje (sorta, ne vrsta); date re-lock blokirajući; glasan pad
link-a bloka + „Izgubljeni" uključuje prazan link; storno filter u
`ExistingBlokCena`/`BuildFirstBlokCena`; kooperant free-text disambiguacija kod >1
pogotka (AUD-028; FM-0007 #2/#3/#5, FM-0009 #4/#5, FM-0008 #1). **Regression:**
`RunBusinessFlowProSuite` + ručni unos otkupa sa panelom blokova.

### RF-12 — Palete [Wave 2 · S/M]
**Fajlovi:** `frmPalete.frm`, `modPaletniList.bas`.
**Obim:** UI validacija „bar jedna vrsta pakovanja" (kutije ILI kese); `Preradjeno`
guard na ulazu Reassign/Detach; sledljivost lista označena kao „mogući izvori"
(AUD-029; FM-0017 #1, FM-0016 #1/#2). **Regression:** `RunPaleteTestSuite`
(modTestPalete).

### RF-13 — Infra lifecycle [Wave 1/2 · M]
**Fajlovi:** `clsTransaction.cls`, `modJournaling.bas`, `modDataAccess.bas`,
`modParse.bas`, `ThisWorkbook.doccls`, `modMain.bas`, `modConfig.bas`,
`modMonitoring.bas`, `modStanicaLock.bas`, `modSetup.bas`.
**Obim:** `RollbackTx` per-tabela trap + garantovan `CleanUp` + `Class_Terminate`;
journal today-vs-today + `UpdateCell` journal + storno-marker na rollback; `AppendRow`
fantomski red; `TryParseDateValue` opseg+round-trip; startup trio (Err capture,
fail-soft backup, `FlushNow` na close); `GetConfigValue` trim; `TBL_CONFIG` van
required listi; monitoring `ActiveWorkbook` fallback; `BulkPushPendingForStanica` →
`UpdateCell` (AUD-004/006/007/017/018). **Regression:** restart aplikacije + namerno
izazvan rollback (checklista daje korake).

### RF-14 — MasterSync / JSON [Wave 1 · M]
**Fajlovi:** `modGoogleSheets.bas`, `modMasterSync.bas`.
**Obim:** JSON čitanje (strip samo van navodnika, `\"`, `\uXXXX`/`\t` dekod);
`ImportOtkupFromPWA_TX` → alias `_Core(False)`; `FindVOZSheets` paginacija
(AUD-001/002 + AUD-018 deo). **Regression:** pun sync ciklus sa test vrednostima
koje sadrže zapete, navodnike i dijakritiku (checklista ih navodi).

### RF-15…RF-19 — Wave 3 konsolidacije [P2 · po jedna sesija]
RF-15 HTTP helperi → `modHttpUtils` + 4 MasterSync call site-a; RF-16
`modBankaParseUtils` (~10 helpera, parseri netaknuti); RF-17 `BrutoUNeto` (6 klonova)
+ selidba `SaveOMUlaz_TX` u `modDokumenta`; RF-18 `NzBlank` + case-insensitive
storno compare; RF-19 `modTestHarness` + `LastRunFailedCount()` + self-update
`files_count`. Detalji: ROADMAP §10.3.

### RF-20 — Wave 4 bezbednost/proces [koordinisano — planirati posebno]
PIN hash + JMBG (VBA+GAS/PWA, migracioni prozor); `saveParcelPolygon` (KI-001);
sređivanje dokumentacije (AR/CL verzije, `instructions/` istorijski, CLAUDE.md
reference); `VBA_SRC_PATH` iz LocalConfig. Detalji: ROADMAP §10.4.

---

## 3b. Paketi iz delte v85 (SEF, startup-auth, sync, self-update, cenovnik)

> Detalji: `docs/AUDIT_FM_TRIJAZA.md` DEO II + `KNOWN_ISSUES.md` §8.4 (AUD-030…039).
> Sidra su `f6313dc`/`a0bc9e2` — pre rada re-bazirati na svež `main` (v2.24.0+).

### RF-21 — SEF correctness [P0/P1 · M]
**Fajlovi:** `modSEFClient.bas`, `modSEFValidator.bas`, `modSEFMapper.bas`,
`modSEFPersistance.bas`.
**Obim:** (P0) 409 izdvojiti iz REJECTED → `apiStatus="CONFLICT"` → TECH_FAILED/manual
(AUD-030); stornirana faktura sendable — `Stornirano` guard u `ValidateFakturaForSEF` +
filter u `frmSEF` combu; qty/price odvojeni `XmlQuantity`/`XmlUnitPrice` (3+ dec);
DueDate ≥ IssueDate uz force-today; `HasSuccessfulSEFSubmission` EH → re-raise
(fail-closed); resubmit čisti stari `SEFDocumentId`; stavke: samo aktivne (Stornirano≠DA,
OsirocenoOd prazno) (AUD-031). **Regression:** `RunSEFTestSuite` + demo SEF: duplicate
submit (409), faktura sa >2 dec, stara faktura (DueDate), storno pa pokušaj slanja.

### RF-22 — SEF UX/lifecycle [P1 · M]
**Fajlovi:** `frmSEF.frm`, `modSEFService.bas`, `modSEFStatusSync.bas`, (+ `modSEFTests`).
**Obim:** posle send-a `frmSEF` bira poruku po `SEFWorkflowState` (ne bezuslovno „Faktura
poslata"); `Test_Cancel/Storno…_TX` iz servisa → `modSEFTests` (ili SEF_ENV guard);
blank/unknown status → `UNKNOWN_STATUS` + manual review (ne SENT); `cmbFaktura_Change` →
`ClearSEFInfo`; recovery/refresh vraćaju rezultat, `frmSEF` ga prikazuje; batch summary
(found/recovered/failed). (AUD-032). **Regression:** ručno slanje odbijene fakture (poruka),
recovery zaglavljenog SENDING.

### RF-23 — Startup + authorization [P1 · S/M]
**Fajlovi:** `ThisWorkbook.doccls`, `modTrial.bas`/`modLicense.bas`, `modMaticniLookups.bas`,
`modAdmin.bas`, `modPodesavanja.bas`, `frmOtkupAPP.frm`, `frmStammdaten.frm`.
**Obim:** `Workbook_Open` → `If AccessWasDenied() Then Exit Sub` (pre `STARTUP_SUCCESS`)
(AUD-034); `MozeAdministraciju` guard: proširiti u `MaticniMenu_OnClick` na Admin/Podešavanja
+ vrh `BuildAdminPanel`/`AdminPanel_OnClick`/`BuildConfigEditor`/`ShowConfigSheet` (AUD-033);
`btnBanka_Click` auth guard pre importa (obrazac iz `btnSyncPWA_Click`) (AUD-034); PasswordChar
za „secret" polja u Podešavanjima; signal pri plaintext PIN fallbacku. **Regression:** login kao
ne-admin → probati Matične → Admin panel mora biti blokiran; deny licence → app se zatvara.

### RF-24 — Self-update hardening [P1 · M]
**Fajlovi:** `modSelfUpdate.bas`, `modRelease.bas`, `modBuildGuard.bas`, `modBuildInfo` gen.
**Obim:** faza 2 pokriva i failed `.frm` (ili ga ne Remove-uje u fazi 1); manifest `files_count`
+ abort na nepotpun download; `PublishReleaseToDrive` guard: placeholder/`+dirty` deny +
disk↔workbook `BUILD_SHA` cross-check; `AssertBlankBuild` skenira plain-range logove
(`SETUP_LOG`, test logovi) (AUD-035, AUD-037). **Regression:** simulirati prekinut download;
publish sa dirty radnim direktorijumom mora pasti.

### RF-25 — Sync/IO hardening [P2 · M]
**Fajlovi:** `modGoogleSyncOrchestrator.bas`, `modGoogleSheets.bas`, `modDrive.bas`,
`modStammdatenSync.bas`.
**Obim:** `SetPWAMasterSyncLock` → RMW obrazac (ne full-tab overwrite koji briše
`STANICA_LOCK_*`); dva rename-a u jedan `batchUpdate` (atomski swap); `ReadSheetData`
EMPTY≠ERROR (ByRef ok); `DriveFindInFolder` error≠not-found; empty-source cloud-wipe guard;
samostalni Parcele export → geo pull gate. (AUD-038). **Regression:** pun sync sa praznim
lokalom (ne sme obrisati cloud), paralelni station lock očuvan.

### RF-26 — Cenovnik + E2E gate [P1/P2 · S/M]
**Fajlovi:** `frmOtkup.frm`, `frmDokumenta.frm`, `modCenovnik.bas`, `modE2EReleaseGate.bas`,
`modBusinessFlowProTests.bas`.
**Obim:** stale auto-cena — očistiti polje pre lookup-a, prazno na 0 (AUD-036); `GetVazecaCena`
`Optional asOfDate` + `Datum` u schema guardu + UI datum fallback → greška; E2E gate: Boolean
`Core` po suite-u, `E2E_Pass` samo na `m_Failed=0`, ILI deprecirati modul; environment guard u
shipped test suite-ovima (AUD-039). **Regression:** unos otkupa za dva proizvoda uzastopno
(cena se ne prenosi); E2E gate mora prijaviti FAIL kad suite padne.

---

## 3c. Paketi iz delte v142 (agrohemija, sync-integritet, dijagnostika, sledljivost)

> Detalji: `docs/AUDIT_FM_TRIJAZA.md` DEO III + `KNOWN_ISSUES.md` §8.6 (AUD-040…048).
> Sidro je `origin/main` v2.24.0 (`9fd7087`) — pre rada re-bazirati na svež `main`.

### RF-27 — Agrohemija cena + validacija [P1 · S/M]
**Fajlovi:** `frmAgrohemija.frm`, `modAgrohemija.bas` (+ `modAgrohemijaTests.bas`, `modPoruke.bas`, `modJournaling.bas` test-mode).
**Obim:** izlaz prosleđuje `m_KorpaIzlaz(i).cena` kao `overrideCena` u `SaveMagacin` (simetrično
sa ulazom `:843`) → knjižena cena = snapshot korpe (AUD-040); `modAgrohemija` zahteva validnu
cenu > 0 za realne artikle (osim `ART_POCETNI_DUG`) umesto tihog `Cena=0/Vrednost=0`; referencijalne
provere u `ValidateMagacinInput` (postoji/aktivan artikal/koop/parcela↔koop). **Regression:** izlaz
artikla čija master cena ≠ cena u korpi → `tblMagacin` red mora nositi cenu iz korpe; izlaz sa
nenumeričkom cenom mora pasti, ne upisati 0.
**Status (grana `claude/rf-27-agrohemija-cena`, pre-merge):** ✅ izlaz `overrideCena`; ✅ fail-closed
cena ≤ 0 (`SaveMagacinCore` diže typed grešku, `SaveMagacin` ostaje back-compat omotač → operater
vidi tačan razlog, ne generički 4301); ✅ referencijalno: artikal/koop postoje, **parcela↔koop
implementirana** (`;`-lista, svaka parcela postoji + pripada koopu + aktivna via `COL_PAR_AKTIVNA`;
`PRACENJE_PARCELA` ON → parcela obavezna, OFF → prazna dozvoljena; `ART_POCETNI_DUG` izuzet);
✅ nova zero-value ULAZ staza uz `allowZeroValue` (izlaz strog). „Aktivan" za `tblArtikli`/`tblKooperanti`
= N/A (nema kolone u šemi). Testovi: `modAgrohemijaTests.RunAgrohemijaSmokeSuite` (izolovano:
dev-guard + `modJournaling` test-mode + TX rollback, bez traga). AUD-040 zatvoren; AUD-049 povučen
(parcela više nije odložena).

### RF-28 — MasterSync integritet delte [P1 · M] (koordinisati sa RF-14)
**Fajlovi:** `modMasterSync.bas`, `modBrojevi.bas`, `modMalina.bas`, `modAutoHladnjaca.bas`.
**Obim:** `GenerateBrojPrijemnice` EH → `""` (ne `1/ddmmyy`) (AUD-041); `GenerateBrojZbirne`
delegira `SuggestNextBroj(KIND_ZBR,…)` umesto row-count (AUD-041); `TryUpdateVozacID` vraća True
samo ako write uspe (`RequireUpdateCell`) (AUD-042); strict datum parse → `SyncError` (ne današnji)
na oba puta (AUD-042); poison spreadsheet — temp-ime/rename ili trash na fail header-a (AUD-042);
auto-otpremnica ključ + `VrstaVoca|SortaVoca|Cena|TipAmbalaze`; VOZ link membership + „prazno ILI
isto" guard (AUD-043); canonical `IsManagedStationMirror` + re-raise u Ensure EH (AUD-046).
**Napomena:** deli `modMasterSync` sa RF-14 (JSON/paginacija) — raditi u istoj sesiji ili strogo
serijski uz re-bazu. **Regression:** sync batch sa 2 ista `BrojZbirne`, nevalidnim datumom, i
stanicom bez `tblVozaci` para — svaki mora dati `SyncError`, ne tihi upis.

### RF-29 — Integritet/health dijagnostika [P1/P2 · M]
**Fajlovi:** `modIntegritet.bas`, `modProductionHealthCheck.bas`.
**Obim:** `WriteErr` uvodi `ErrorCount`; overlay/MsgBox prikazuje INCOMPLETE kad errors>0; typed
`IntegrityRunResult` (Empty ≠ PASS) (AUD-044); health SEF lista koristi `WF_SEF_*` konstante +
state matrica (ukloniti nepostojeći `SEF_CANCELLED`); parent OK uslovljen child delta-brojačima
(`:951`, `:928`) (AUD-047). **Regression:** namerno pokvaren red (schema drift) → integritet mora
prijaviti GRESKA + broj, ne „0 neusklađenih"; health sa child FAIL ne sme dati parent OK.

### RF-30 — Sledljivost trace + sitni lifecycle [P1/P2 · S/M]
**Fajlovi:** `modSledljivost.bas`, `frmSledljivost.frm`, `modStornoWarm.bas`.
**Obim:** `TraceByZbirna` direktan `BrojZbirne` prolaz kroz `tblOtkup` + `UCase$(Trim$())`+
`vbTextCompare` (isto kao auto-link); typed trace rezultat (`IsComplete`); forma/PDF oznaka
„NEPOTPUN TRACE" (AUD-045); `modStornoWarm` flag tek po uspehu `OnTime` (`m_warmScheduled`),
`LogErr` na cancel fail (AUD-048). **Regression:** otkup sa `BrojZbirne` a praznim `OtpremnicaID`
mora ući u trag; PDF nepotpunog traga mora biti obeležen.

> **Banka (RF-16):** v142 FM-0128..0132 potvrđuju deljene rupe parsera (datum shape-only,
> račun bez normalizacije, poziv preuzak, 0/0-sa-računom) → hrane `modBankaParseUtils` (RF-16):
> kalendarska validacija datuma, opc. mod-97 računa, prošireni poziv obrazac, `zad>0 Xor odo>0`.

## 4. Status

| Paket | Naziv | Status | Grana | Napomena |
|---|---|---|---|---|
| RF-01 | Brisanje balasta | ✅ merged | PR #147 | M0 · AUD-016 (deo) |
| RF-02 | modNovac guardovi | ✅ merged | PR #148 | M1 · AUD-003 (SaveNovac+AddCena) + AUD-010 |
| RF-03 | Storno lanac | ✅ merged | PR #167 | M3 · AUD-020/021 + AUD-049 (storno izvoda) + keš/virman; review OK, follow-up AUD-050/051 |
| RF-04 | AutoHladnjaca | ⬜ | — | |
| RF-05 | frmDokumenta set | 🟢 PR | `claude/rf-05-frmdokumenta-fixes-63yqjp` | M3 · AUD-009 + AUD-022 + deo AUD-003; uz to nova `GeneracijaID` kolona (schema) i guard protiv storna po nejedinstvenom broju (AUD-052 novo). BFP 276/276, Storno 181/181 |
| RF-06 | modIzvestaj brojke | 🟢 PR #175 | `claude/rf-06-izvestaj-brojke-wquclc` | M5 · AUD-023 zatvoren (FM-0028 #1/#3/#5/#6/#9/#10/#12/#13/#14) + posledica AUD-052 u report sloju. **`Compile` čist, `RunIzvestajTests` 100%** (uklj. 3 e2e nad tabelama). Ostaje uporedni pregled izveštaja pre/posle — brojke se namerno menjaju |
| RF-07 | frmIzvestaj + revers | ✅ merged PR #176 | `claude/rf-07-izvestaj-freshness-u0jy43` | M5 · AUD-024 + AUD-012 zatvoreni, AUD-027 delimično (samo cross-tab print; reprint stornirani + `FillFakturaSablon` `.UnMerge` ostaju RF-08). Novi seam-ovi `IzvestajTabDostupan`/`IzvestajEntitetKod`/`ReversRedPripada` + 3 test grupe u `RunIzvestajTests`. Freshness/CleanFail/tab-meni = operater-smoke. **7 novih `Poruka()` ključeva → `EnsurePoruke` obavezan** |
| RF-08 | Faktura + štampa | 🟢 grana | `claude/rf-08-faktura-stampa-g5h1c5` | **M5 ✅ KOMPLETAN.** AUD-011 zatvoren (FM-0034 #1/#2/#3: vlasništvo prijemnice u `CreateFaktura`, fail-closed `rows.count > 1`, `CreateFaktura` → `Private` — caller check potvrdio da svi pozivi već idu kroz `_TX`) + AUD-027 zatvoren u celosti (FM-0031 #3 reprint storniranog otkupa preko nove **fail-closed kapije** `RequireOtkupAktivanZaStampu` — puca na nedostajuću kolonu, nepostojeći i dupliran `OtkupID`, ne samo na storno; #19 `FillFakturaSablon` `.UnMerge` na **dinamičnom** opsegu `max(nStavke + 4, SablonLastContentRow)` — po SADRŽAJU (`Range.Find` nad `A:F`), ne po `UsedRange` (broji i prazne formatirane ćelije), ne `ws.cells` i ne fiksnih 80 redova — šablon je perzistentan i ima namerne merge-ove u zaglavlju). `RunFakturaSmokeSuite` postao **tvrd gate** (uz `On Error GoTo 0` pre gate-a) + 4 nova testa, uklj. automatizovan **81 → 82 stavke** i assert da formatirana prazna ćelija ne širi granicu čišćenja. Merge test malog obima (3 posle 1) = operater-smoke. **2 nova `Poruka()` ključa → `EnsurePoruke` obavezan** |
| RF-09 | Banka import/map | 🟢 grana | `claude/rf-09-banka-import-mapiranje-87nv8k` | **M6, drugi paket.** AUD-014 + AUD-025 zatvoreni, **P0 AUD-007 preuzet ovde** (bio siroče — nije bio zaveden ni na jednom RF-u). (a) `TryParseDateValue` opseg + round-trip (`DateSerial` je nemoguć datum tiho prelivao u sledeći mesec) + **Level 0 kapija datuma u `ParseBankaIzvodForImport`** (datum izvoda i svake transakcije se validiraju pre staging-a; nevalidan → uvoz tog PDF-a pada, fajl u Error folder); (b) `IsDuplicateBankaImport` dobio obavezan `brojRacuna` u obe grane ključa — ista transakcija na drugom računu firme se više ne gubi kao „duplikat"; (c) `GetOtkupCandidatesForKooperantBlock`: bafer veličine ulaza + `MAX_BLOK_KANDIDATA` → tipizirana `ERR_BMAP_MANUAL_REQUIRED`, a nov `AutoMapBankaImportRowBatch` guta **samo taj** kod (diže se pre ijednog upisa) → jedan anomalan blok više ne rollback-uje ceo `AutoMapAll`; (d) nov `RequireBimSmer` u `Map*` (Kupac=uplata, Kooperant/blok=isplata, OM=čist smer) — pogrešan smer je iz UI-a bio dostupan; (e) `frmBankaImport` na `_Activate` više ne knjiži — samo prebroji (`CountStrongKeyReadyBankaImport`, read-only) i ponudi **runtime dugme „Mapiraj jake ključeve (N)"** (`clsUiSink`, `.frx` netaknut); maska `On Error Resume Next` uklonjena; (f) `cmbFaktura` oživljena (otvorene fakture tog kupca) + potvrda „bez fakture = avans" + `RequireFakturaZaKupca` na upisu; (g) `EffectiveBlockNo` + sekcija „RUČNO" u preview-u — prikaz i akcija čitaju iste kontrole. **Nov `RunBankaImportTestSuite`** (modTestBanka, 4 tvrde kapije). Bez novih `Poruka()` ključeva |
| RF-10 | Banka export | 🟢 grana | `claude/rf-10-banka-export-uo4pg8` | **M6, treći paket — milestone ✅ kompletan.** AUD-026 zatvoren u celosti. (a) **Klamp zaostalog override-a (FM-0020 #1):** per-blok „Isplatiti" override preživljava reload (namerno — unos se ne gubi), ali ga `PruneStaleOverrides` nije poredio ni sa čim, nego je brisao samo unos za blok koji je iz liste nestao; posle delimične isplate, vezanog avansa ili storna zaostao iznos je mogao biti **veći od otvorenog** i takav ući i u prikaz i u CSV nalog. Logika je izmeštena u nov `modBankaExportPregled.ClampOverridesToOpen(overrideDict, blokovi)` (seam van forme, uz `BuildOpenAmountDict` kao zajednički indeks `OtkupID → OtvorenIznos` koji koriste i klamp i revalidacija): nestao/zatvoren blok gubi override, veći se spušta na otvoreno, a **manji ostaje netaknut** — delimična isplata je operaterska odluka, ne stale podatak. `PruneStaleOverrides` je sada `Function` koja vraća broj usklađenih unosa, pa `LoadBlokovi` to javlja u statusu (tiho spuštanje iznosa bi operater promašio). (b) **Finalna saldo kapija (FM-0020 #2 / FM-0021 #1):** `GenerisiNalogeCSV` je pred upis proveravao **samo** `HasTekuciRacun` i `IsplatitiIznos > 0` — ništa o saldu — pa je promena podataka između prikaza i klika na „Generiši" mogla da završi kao preplata u nalogu banci. Sada se pred upis čita **svež** saldo (`BuildBlokIsplataList()`) i svaki nalog poredi kroz nov `ValidateNalogSaldo`; preplata → **fajl se ne piše**, razlog (koji blok, traži vs otvoreno) ide kroz nov `Optional ByRef outOdbijeno`, forma ga prikaže i odmah osveži pregled. **Fail-closed na tri načina:** blok kog više nema među otvorenima = otvoreno 0; greška pri očitavanju završi u `EH`, dakle takođe bez fajla; a provera gleda **tačno** one blokove koje writer i upisuje (sa TR, iznos > 0), da blok koji uopšte ne ide u fajl ne obori generisanje. Poređenje je u **cent-domenu** preko postojećeg `modNovac.ZaokruziNovac` (half-up), **bez epsilon tolerancije** — raniji prag `> otvoreno + 0.01` je propuštao preplatu od punog centa (`600.01` nije veće od `600.00 + 0.01`), a baš taj cent banka isplati; isto pravilo važi za klamp, finalnu kapiju, operaterski unos (`txtIsplatiti_Exit`) i `CsvIznos` (normalizuje PRE `Format$`), pa se ono što kapija odobri i ono što writer upiše ne mogu raziići. NE uvoditi novu granicu ni nov helper. (c) **Brojanje stvarno primenjenih avansa:** `ApplyAvansToOtkup_TX` vraća `True` i za no-op (uspela TX bez ijednog proknjiženog dinara), pa je batch „Primeni avans (sel.)" prijavljivao naduvan broj; `PrimeniAvansTX` sada čita primenjeni iznos iz `ByRef` parametra koji je **RF-02/AUD-010 već uveo** (reuse, ne nov ugovor) — batch razdvaja primenjeno / bez promene / greška uz zbir, po-blok akcija ne prijavljuje uspeh kad ništa nije knjiženo, a override se skida samo kad se otvoreno bloka zaista promenilo. **Testovi:** `modTestBanka` proširen sa RF-09 na celu banku — T12 (klamp: 1000 na otvoreno 600 → 600, nestao i zatvoren blok → brisanje, manji override netaknut, drugi prolaz no-op) i T13 (nalog preko otvorenog odbijen **sa imenom bloka**, nestao blok fail-closed, blok bez TR ne obara naloge, razlika ispod tolerancije prolazi, bez mape otvorenih odbija); čisti seam testovi bez tabela, van transakcije kao T01. **Van obima (svesno):** PDF specifikacija isplata nema istu kapiju (prikazni dokument, ne nalog banci); export ledger/rezervacija bloka (FM-0021 #3) i agregacija klasa istog `BrDok` (FM-0021 #2) ostaju P2 u katalogu. **Review runda 1 (pre merge-a):** (R1) poredjenje iznosa je bilo na epsilon pragu `> otvoreno + 0.01` (preuzet iz zatecene provere unosa) — taj prag propusta preplatu od **punog centa** (`600.01` nije vece od `600.00 + 0.01`), a bas taj cent banka isplati; sva tri mesta (unos u formi, klamp, finalna kapija) sada normalizuju oba iznosa postojecim `ZaokruziNovac` (half-up) i porede **tacno**, `CsvIznos` normalizuje **pre** `Format$` (inace bi writer i validator mogli da se raziidju), a provera „iznos > 0" ide nad zaokruzenom vrednoscu da sub-cent unos ne prodje validaciju pa tiho ispadne iz fajla. Uvedene granicne provere `600.00/600.01/600.005/600.004`. (R2) test je gadjao samo validator, pa bi uklanjanje njegovog poziva proslo nezapazeno — sadrzaj CSV-a se sada gradi iskljucivo kroz nov `BuildNalogCsvPayload` koji **u sebi nosi** kapiju (jedina putanja do `WriteAllTextUtf8`), a nov T14 tvrdi da je payload pri preplati prazan. (R3) `Err.description` se u svim `EH` granama ovog ekrana hvata **pre** `LogErr` (BUG-1/AUD-054) — ranije je operater na pad generisanja naloga mogao dobiti praznu poruku umesto razloga. (R4) sub-cent ostatak je mogao da napravi nalog na **nula dinara**: `GetOpenOtkupi` racuna otvoreno sirovo (`kolicina * cena - isplaceno`, bez `ZaokruziNovac`), pa je legitiman ostatak tipa `0.004` RSD prolazio kroz sva tri praga `> 0` (forma `CollectIsplataBlokovi`, validator, payload writer), a `CsvIznos` bi ga upisao kao `;0.00;` — nije preplata, ali je nevalidan nalog koji banci moze da obori uvoz CELOG paketa. Iznos se sada normalizuje **pre** praga na sva tri mesta (forma pamti zaokruzen `IsplatitiIznos`, validator i writer koriste isti `trazeno = ZaokruziNovac(...)`), pa takav blok jednostavno ne dobija nalog. T14 dopunjen slucajem `regularan nalog + blok 0.004` → tacno jedan data red i payload bez `;0.00;`. (R5) saldo-mapa nije bila fail-closed na **korumpiranom kanonskom ključu**: `BuildOpenAmountDict` je duple `OtkupID`-eve pisala assignment-om (svesna odluka iz prve verzije „da korumpiran podatak ne obori pregled"), a `GetOpenOtkupi` ne agregira po `OtkupID` nego vraća red po red — pa bi kod dva reda sa istim `OtkupID`, **različitim** `BrojDokumenta` i različitim otvorenim iznosom poslednji upis odlučio saldo, i otvoreno bloka B bi odobrilo nalog za blok A (`traženo 1.000 <= 2.000` prolazi iako je na A otvoreno 500). Na finansijskoj kapiji je to fail-**open**, pa je vraćeno na tvrd pad: dupli i prazan `OtkupID` sada dižu tipizirane greške (`ERR_ISPLATA_DUPLI_OTKUPID` / `ERR_ISPLATA_PRAZAN_OTKUPID`). Posledica je namerna na obe strane — pregled prijavi razlog u statusu umesto da „nekako" izabere koji duplikat važi, a `GenerisiNalogeCSV` padne u svom `EH` **pre** nego što payload uopšte nastane (mapa se gradi kao argument), pa nema fajla za banku; razlog se sada vraća kroz `outOdbijeno` umesto generičnog „pogledajte log". Duplikat inače hvata i `RunProductionHealthCheck`, ali health-check nije preduslov svakog klika na „Generiši naloge". Nov T15 gađa konfliktne vrednosti (500 vs 2.000 pod istim `OtkupID`), propagaciju kroz klamp i prazan ključ; zdrava lista mora i dalje da prođe. (R6) **cross-flow rupa iste invarijante:** R5 je zatvorio CSV, ali ne i drugu mutacionu akciju istog ekrana. `ClampOverridesToOpen` je izlazio pre izgradnje mape kad override-a nema (`overrideDict.count = 0`) — a to je stanje pri **prvom** otvaranju forme — pa strogа provera kanonskog ključa nije ni radila i korumpiran `OtkupID` je mogao normalno da stoji na ekranu; mapa se sada gradi **pre svih ranih izlaza**. Važnije: `modNovac.ApplyAvansToOtkup` je target red čitao kao `FindRows(...)(1)`, dakle **prvi red pobeđuje**, bez `count > 1` provere — kod dupliranog `OtkupID`-a vrednost i vlasnik se čitaju sa jednog reda, a avans se u ledgeru vezuje samo za dvosmislen `OtkupID`. Guard je zato dodat u **core** (`vbObjectError + 1045`, pre svakog čitanja target reda i pre ijednog upisa), ne u formu — UI nikad ne sme biti jedina bezbednosna granica, a `RunProductionHealthCheck` nije preduslov klika na „Primeni avans". Nov **T16** je tabelarni (ne dictionary) i gađa slučaj koji je ranije **tiho** prolazio — **isti** kooperant na oba duplirana reda (različit kooperant je već obarao zatečeni owner guard AUD-010, pa se i on proverava da ne oslabi): akcija mora biti odbijena, `appliedAmount = 0`, nijedan `tblNovac` red ne sme dobiti sporan `OtkupID`, avans ostaje nevezan, oba `tblOtkup` reda netaknuta — uz kontrolu da jedinstven `OtkupID` i dalje prolazi. Provereno je i da duplikat ne nastaje ni u jednom legitimnom toku: `tblOtkup` se puni samo na dva mesta (`modOtkup`, `modMasterSync`), oba preko `GetNextID`, a storno postojeće redove samo označava. (R7) **najopasniji sloj iste invarijante — pogrešan primalac novca:** stroga saldo-mapa iz R5 vidi samo **otvorenu** kolekciju, a identitet kooperanta je dolazio iz `BuildLookupDict(TBL_OTKUP, ...)` — „prvi pojav pobeđuje" nad **sirovom** tabelom, dok `GetOpenOtkupi` preskače stornirane i već isplaćene redove. Kod dupliranog `OtkupID`-a te dve semantike se razilaze: red A (kooperant K-A) isplaćen i zato nevidljiv u otvorenoj listi, red B (kooperant K-B) otvoren — u otvorenoj kolekciji postoji **jedan** `OtkupID`, pa duplikat-kapija nema šta da prijavi, ali vlasnik dolazi sa reda A i blok B dobija naziv i **tekući račun kooperanta A**. Nalog bi otišao **pogrešnom primaocu**, što je teže od pogrešnog prikaza. Zato se identitet primaoca više ne rešava „prvim pogotkom": nov `BuildOtkupOwnerIndex` u jednom prolazu kroz sirovu tabelu pamti vlasnika **samo za jednoznačne** `OtkupID`-eve i zasebno evidentira duplikate, a `BuildBlokIsplataList` pada (`ERR_ISPLATA_DUPLI_OTKUPID`) čim je duplirani ID **otvoren** — provera ide pre view-filtera, da se ekran ponaša isto kao finalna kapija koja listu gradi bez filtera. Istorijski duplikat kome nijedan red nije otvoren namerno **ne** obara ekran. Nov T17 daje kooperantima očigledno različite račune i dokazuje da nikakav payload ne nastane, uz kontrolu da zatvoren duplikat prolazi; T17 se u suite-u pokreće **pre** T16, jer T16 namerno ostavlja dva otvorena reda sa istim `OtkupID`. Uz to: neuspeo rebuild više ne ostavlja **prethodnu** listu na ekranu (`LoadBlokovi` EH čisti prikaz) — writeri jesu fail-closed, ali prikaz ne sme da tvrdi više od njih. (R8) **poslednji tihi gubitak obaveze:** RF-10 je uveo `ERR_ISPLATA_PRAZAN_OTKUPID`, ali taj guard nikad nije mogao da opali — `BuildOtkupOwnerIndex` prazan `OtkupID` preskače (ispravno, to je samo indeks), pa vlasnik nije nađen i red je završavao u zatečenom tihom `GoTo NextRow`: **otvorena obaveza je nestajala sa ekrana bez ijedne poruke**, dok bi ostali nalozi uredno otišli u fajl. Isto je važilo za otvoren blok sa praznim `KooperantID` (FM-0021 #5). To je suprotno pravilu koje ovaj paket uvodi, pa je identitet sada fail-closed po redosledu: `OtkupID` postoji → jedinstven je → kooperant postoji → tek onda TR/naziv/filteri. Prazan kooperant nije legitimno stanje — `SaveOtkupMulti_TX` ga odbija pri upisu (greška 1810), a `OtkupID` dolazi iz `GetNextID`; dakle nijedan zdrav tok ne može da napravi takav red. **FM-0021 #5 je time zatvoren strože nego što je katalog predlagao** (predlog je bio „prebroj preskočene i prikaži”, ovde je tvrd pad). Nova `ERR_ISPLATA_PRAZAN_KOOPERANTID`. Novi T18/T19 (tabelarni): otvoren blok bez `OtkupID` odnosno bez kooperanta mora da padne tipizirano i da imenuje dokument, uz kontrolu da **zatvoren** red bez `OtkupID` ne obara pregled (granica iz R7: korupcija koja ne može proizvesti nalog ne blokira ekran). Uz to su T16–T19 prebačeni u **izolovane transakcije** (`BeginIsolatedTx`) — svi čitaju CELU tabelu, pa je seed jednog obarao kontrolne slučajeve drugog; izolacija je uklonila i zavisnost od redosleda pokretanja. **Bez novih `Poruka()` ključeva** |
| RF-11 | Otkup UI | ⬜ | — | |
| RF-12 | Palete | ⬜ | — | |
| RF-13 | Infra lifecycle | ⬜ | — | |
| RF-14 | MasterSync/JSON | ⬜ | — | **koordinisati sa RF-28 (isti fajl)** |
| RF-15–19 | Konsolidacije | ⬜ | — | posle RF-14; RF-16 hrani v142 banka delta |
| RF-20 | Bezbednost/proces | ⬜ | — | planirati posebno |
| RF-21 | SEF correctness | ✅ merged | PR #152 | M2 · AUD-030 (P0 409) + AUD-031 |
| RF-22 | SEF UX/lifecycle | 🟢 grana | `claude/rf-22-sef-ux-lifecycle-kzzzvn` | **M6, prvi paket.** AUD-032 zatvoren u celosti: (a) ugovor o ishodu slanja — `IsSuccessfulSEFSendState`/`SEFSendFailureErrNumber`/`SEFSendOutcomeMessage`, neuspeh diže tipiziranu grešku (posle commit-a i `On Error GoTo 0`) umesto da vrati SubmissionID; (b) prazan/nepoznat status → `UNKNOWN_STATUS` + ručna provera, nikad tiho SENT (`IsKnownSEFRefreshStatus`, `SEFUnknownStatusTargetState`); (c) `RefreshSEFStatus_TX` vraća stvaran ishod, recovery upisuje „Recovered" samo kad faktura zaista izađe iz `SEF_SENDING` — plus dva uzroka večne petlje (remote STORNO dok je lokalno SENDING; pad refresh-a je ciljao zabranjenu tranziciju `SENDING→SYNC_ERROR`); (d) `cmbFaktura_Change` → `ClearSEFInfo` (bez novih form `WithEvents`); (e) `Test_Cancel/Storno…_TX` uklonjeni (gađani ekvivalenti postoje u `modSEFTests`), ostali dev makroi → `Private` (van Alt+F8); (f) batch sažetak (`Found/Recovered/NotRecovered/Failed`). Usput: `SEF_UNKNOWN` izlazne tranzicije (bio slepo crevo) i 1 AUD-054 site u `modSEFStatusSync`. **Nov `RunSEFTestSuite`** — offline tvrd gate (bio referenciran u planu, nije postojao) + 5 seam testova. **Review runda 2:** (R1) adapter za zvanični `SalesInvoiceStatus` — SEF prihvatanje zove `Approved`, ne `Accepted`, pa bi odobrena faktura završavala kao „nepoznat status"; jedan `ClassifySEFExternalStatus` za sve potrošače (+`APPROVED` u storno kapiju i dugme). (R2) čist planer `SEFRefreshTargetState` koji **pita** state machine (`IsSEFTransitionAllowed`) — ponovni pad SEF-a nad `SEF_UNKNOWN`/`SEF_SYNC_ERROR` je pucao na zabranjenu tranziciju, a `UNKNOWN` je bio trajan; `ApplySEFStateOrRefreshOnly` uklonjen. (R3) `IsSEFRecoveryComplete` → whitelist (prazno/`BOGUS` je prolazilo kao oporavak). Orkestracioni test **12 stanja × 8 klasa** našao regresiju u prvoj verziji planera. **Review runda 3:** (R4) `MISTAKE` (greška pri slanju) izvučen iz `TERMINAL` u novu klasu `SEF_CLS_SEND_FAILED` → `SEF_TECH_FAILED` (retry putanja) — ranije je neuspelo slanje postajalo lokalno `SEF_SENT`, bez retry-ja, bez cancel-a i preskočeno u batch-u; (R5) `INFO` (`Paid/OverDue/Archived`) sada izvlači `SENDING`/`UNKNOWN`/`SYNC_ERROR` u `SEF_SENT` — dokazuje da dokument nije više „u slanju", iako ne dokazuje prihvatanje; (R6) telemetrija razdvojila `SEF_STATUS_INFO` od `SEF_STATUS_TERMINAL`. Cancel/storno spiskovi spojeni u `CanCancelSEFStatus`/`CanStornoSEFStatus` (validator + forma isti izvor) uz capability test. **Review runda 4:** (R7) dodate tranzicije `SEF_SENT → SEF_TECH_FAILED` i `SEF_SYNC_ERROR → SEF_TECH_FAILED` — normalna sekvenca (submit ok → `SEF_SENT` → refresh vrati `MISTAKE`) je ostajala „poslata" i batch ju je osvežavao u nedogled; (R8) `CanSendSEFInvoice(workflow, sefStatus)` — sam workflow nije dovoljan, jer `SEF_TECH_FAILED` iz `MISTAKE` ima živ dokument na SEF-u i duplicate guard bi ga odbio: **poslovna odluka je Cancel + ručna provera, ne Retry** (SEF-ov ugovor za ponovni POST istog `requestId` nije proveriv statički — zapisano kao jedna otvorena odluka za vlasnika); (R9) ista kapija rešava i „Retry upaljen posle uspešnog Cancel-a", bez uvođenja `WF_SEF_CANCELLED`. **Review runda 5:** (R10) `CanSendSEFInvoice` je postojanje dokumenta izvodio iz **promenljivog** `SEFStatus`-a — pad refresh-a ga prepiše u `HTTP_ERROR` i „Retry" se ponovo pali; potpis proširen `sefDocumentId`-em i odluka vezana za trajan zapis (status ostaje rezervna odbrana), `REJECTED` sendable samo iz `SEF_READY`; (R11) `SEFSendBlockedNextStep` — poruka više ne savetuje Cancel tamo gde Cancel nije dozvoljen. **Review runda 6:** (R12) resubmit odbijene fakture je bio **mrtav tok** — refresh ne dira submission red, pa je `HasSuccessfulSEFSubmission` obarao baš onaj resubmit koji je `PrepareRejectedInvoiceForResubmit` pripremio; nov `DischargeSEFSubmission_Row` razdužuje **tačno** poslednju `SENT` submisiju (uz proveru vlasništva) u istoj TX, priprema dobila fail-closed post-proveru (`HasSuccessfulSEFSubmission` i dalje `True` → glasan pad, ručna provera), telo pripreme izdvojeno u `_Row` (jer `BeginTx` puca na ugnežđenu TX, pa se tok nije mogao testirati) i dodat **integracioni** test kroz stvarne tabele + stvarni duplicate guard, pod `modJournaling.SetTestModeQuiet` (rollback ne vraća CSV journal). **Review runda 8 (pred merge):** popravljena **sva 22** AUD-054 mesta u `modSEFPersistance`/`modSEFValidator` (writer i duplicate guard nisu propagirali grešku, pa su fail-closed kapije bile deklarativne; ostatak 21); uspešan **storno** sada pomera i lokalni workflow u `WF_SEF_STORNO` preko zajedničkog `ApplySEFExternalOutcome_Row` (pre toga trajna kontradikcija `SEF_SENT` + `STORNO`, koju je učvrstio sam planer), a `Cancelled`/`Deleted` su dokumentovani kao external-terminal-only; priprema resubmita više ne piše interni marker u `SEFStatus`; razduženje proverava i `SEFDocumentId` lineage; nov table-level test `Test_SEFStornoMovesLocalWorkflow`. **Runde 9–10:** ispravljen kontradiktoran assert koji bi oborio suite (`Storno` klasa), dodat `WF_SEF_STORNO` u `IsSEFRecoveryComplete` (recovery je prijavljivao neuspeh za ispravno storniranu fakturu), pokrivena lineage grana razduženja, i dodat `tools/check-sef-asserts.py` — statička provera koja čita produkcioni izvor i hvata testove koji tvrde nešto što kod više ne radi (CI ne pokreće Excel). **7 novih `Poruka()` ključeva → `EnsurePoruke` obavezan** |
| RF-23 | Startup + authorization | ✅ merged | PR #149 | M1 · AUD-033/034 (auth lanac + AccessWasDenied) |
| RF-24 | Self-update hardening | ⬜ | — | |
| RF-25 | Sync/IO hardening | ⬜ | — | |
| RF-26 | Cenovnik + E2E gate | ⬜ | — | |
| RF-27 | Agrohemija cena + validacija | ✅ merged | PR #154 | M2 · AUD-040 + parcela↔koop + typed greške (`SaveMagacinCore`) + zero-ULAZ + smoke suite |
| RF-28 | MasterSync integritet delte | ⬜ | — | spojiti sa RF-14 (isti fajl) |
| RF-29 | Integritet/health dijagnostika | ⬜ | — | |
| RF-30 | Sledljivost trace + lifecycle | ⬜ | — | |

**Redosled:** RF-01 → RF-02 → RF-23 → RF-21 → RF-27 → RF-03* → RF-04* → RF-05 →
RF-14+RF-28 → RF-06 → RF-07 → RF-08 → RF-22 → RF-09 → RF-10 → RF-11 → RF-12 → RF-13 →
RF-26 → RF-29 → RF-30 → RF-24 → RF-25 → RF-15+ → RF-20.
- **RF-27 podignut napred:** agrohemija cena≠knjižena je P1 finansijski/audit nalaz sa jeftinim
  (S) fixom — najbolji odnos vrednost/rizik u v142 delti.
- **RF-14 + RF-28 zajedno:** oba diraju `modMasterSync` (JSON/paginacija vs generatori/writeback);
  raditi u jednoj sesiji da se izbegne dupla re-baza istog fajla.
- **RF-23 i RF-21 podignuti napred:** RF-23 nosi P0-klasu bezbednosti (auth lanac do Admin
  panela) i P1 `AccessWasDenied`; RF-21 nosi jedini nov P0 (SEF 409). Oba pre storna.
- **RF-03*/RF-04* (storno):** OBAVEZNO re-verifikovati protiv `origin/main` (v2.24.0, storno
  PR #134–137 su prepravili `modStornoFlow` +746 linija) — deo nalaza je možda već rešen.
- Redosled se sme menjati — pravilo je samo: serijski, jedan po jedan, re-baza na svež `main`.
