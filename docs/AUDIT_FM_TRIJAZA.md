# Trijaža nalaza iz AgriX Functional Map v35 — stavka po stavka

**Datum:** 2026-07-18
**Izvor nalaza:** `AgriX_Functional_Map` v35 (FM-0002…FM-0034; ukotvljen na commit `a0bc9e2` / vba v2.21.0)
**Metod:** svaka evidentirana rizik-stavka (665 ukupno: 91 Kritičan, 257 Visok, 260 Srednji, 57 Nizak) pojedinačno je ocenjena uvidom u aktuelni kod u `src-vba/` (8 paralelnih verifikacionih prolaza; Kritičan/Visok stavke uz obavezan citat fajl:linija). FM-0001 (`modConfig`) ne sadrži rizik-tabelu, pa nema stavki za trijažu.
**Kalibracija:** aplikacija je single-writer desktop (drugi Excel otvara fajl read-only; PWA ivica ide kroz GAS lock-ove) — „multi-user race / cross-user CAS" tvrdnje ocenjene su u tom kontekstu (realna varijanta: ponovni ulazak u istoj instanci).

## Legenda

**Opravdanost:** **Tačno** (potvrđeno u kodu, naveden dokaz) · **Delimično** (jezgro tačno, formulacija/težina preterana ili uslovljena) · **Netačno** (opovrgnuto uvidom u kod) · **Dizajnersko ograničenje** (tačno, ali svesna odluka — dokumentovati, ne hitno menjati) · **Nije proverivo statički** (traži runtime test ili poslovnu odluku)

**Hitnost:** **P0** (može oštetiti/izgubiti podatke ili novac — odmah) · **P1** (funkcionalni bug vidljiv operateru, uklj. pogrešan izveštaj/saldo/dokument) · **P2** (hardening / tehnički dug) · **P3** (kozmetika/nisko) · **Prihvaćeno** (svesna odluka; samo dokumentovati)

**Napor:** S (do ~1h) · M (do ~1 dan) · L (višednevno/koordinisano)

---

## Blok A — Infrastruktura (FM-0002…FM-0005, 47 stavki)

### FM-0002 — `modDataAccess.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | Delimičan insert: greška posle `ListRows.Add` ostavlja red | Visok | **Tačno** — `ListRows.Add` na :198, upis :204-206, ErrHandler :219-221 bez `newRow.Delete`; fantomski red van TX ostaje | P1 | U ErrHandler best-effort `newRow.Delete` (modDataAccess.bas) | S |
| 2 | `UpdateCell` nema journal / old-new hook | Visok | **Tačno** — `WriteJournalRow` samo u AppendRow (:210); UpdateCell :224-250 bez hooka; journal = crash-recovery CSV (modJournaling.bas:44) → izmene nepokrivene recovery-jem | P2 | Opcioni journal hook u UpdateCell (tbl, red, kolona, old, new) u modJournaling | M |
| 3 | Bypass tvrdnje „jedini pristup tabelama" | Visok | **Tačno** — direktan `DataBodyRange` u 11 modula van sloja (modStanicaLock ×20, modMigracija ×10, modSEFPersistance ×4, modMonitoring ×4…); uglavnom infra sloj | P2 | Dokumentovati legitimne izuzetke (setup/lock/monitoring/migracija); poslovni upisi isključivo kroz modDataAccess | S |
| 4 | `GetNextID` max+1 nebezbedan za paralelizam | Srednji | **Delimično** — mehanizam tačan (:344-382), ali app je single-writer (drugi Excel = read-only); realan ostatak je re-entry u istoj instanci | P2 | Dokumentovati single-writer pretpostavku; GetNextID+AppendRow držati zajedno unutar TX | S |
| 5 | `StampRowAudit` silent failure (`On Error Resume Next`) | Srednji | **Tačno** — :259; upis prolazi bez audit vrednosti bez ikakvog signala | P2 | Log/`Debug.Print` pri grešci stampa (bez prekida upisa) | S |
| 6 | Persistence zna za KPI dashboard (`gKpiDirty`) | Srednji | **Tačno** — :214-215; svesna performance sprega | P3 | Ostaviti; callback registar tek uz „rule of three" | S |
| 7 | `GetTable` linearno skenira sve sheetove | Srednji | **Tačno** — :73-86; van cache prozora svaki poziv skenira; nije izmeren problem | P3 | Modul-level dictionary ime→ListObject sa invalidacijom | S |
| 8 | Exact-match bez normalizacije (CheckDuplicate/lookup) | Srednji | **Tačno** — :419 i :460 case-sensitive bez Trim; i `GetTable` je case-sensitive (:79) vs `FindListObject` insensitive (modSetup:1564) | P2 | `Trim$` + `StrComp vbTextCompare` u CheckDuplicate; pre šire promene popis pozivalaca | S |
| 9 | Bez column guardova (`GetLookupList`, `GetNextID`) | Nizak | **Tačno** — GetNextID ne proverava `colIdx=0` (:349→:359 subscript greška); GetLookupList isto (:521→:538, filterIdx :543) | P2 | `If colIdx = 0 Then Exit Function` guardovi | S |
| 10 | `Application.Transpose` na velikim kolonama | Nizak | **Tačno** — :178; poznata ograničenja tek na ekstremnim veličinama | P3 | Ručna petlja umesto Transpose ako tabele porastu | S |

Bilans: 9 Tačno / 1 Delimično / 0 Netačno / 0 Dizajnersko ograničenje.

### FM-0003 — `clsTransaction.cls`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | Zaboravljena tabela u snapshot listi ruši atomicitet | Kritičan | **Tačno** (strukturno) — rollback vraća samo `mSnapshots` (:80-83); teret je manuelan i širok: 323 BeginTx/AddTableSnapshot poziva u 27 fajlova; konkretan propust ovde nije dokazan | P1 | Audit mapa po *_TX wrapperu (menja vs snapshotuje); opcioni auto-snapshot helper pre prvog upisa | M |
| 2 | `RollbackTx` bez error-safe finalizacije (`CleanUp`) | Visok | **Tačno** — :77-87 bez handlera; `RestoreTable` može `Err.Raise` (:110-111, ili GetTable=Nothing→err 91) → `CleanUp` :85 preskočen: events off, manual calc, delimičan rollback | P1 | Handler u petlji (nastavi sledeću tabelu, loguj) + garantovan `CleanUp` na izlazu | S |
| 3 | `Value2` snapshot gubi formule | Visok | **Delimično** — mehanizam tačan (:51, :141), ali nijedan modul ne upisuje `.Formula` (grep=0): tabele su data-only, praktičan rizik mali, „Visok" preteran | P3 | Dokumentovati pravilo „bez formula u data kolonama"; bez izmene koda | S |
| 4 | Spoljni side effecti (PDF/HTTP/GAS) van rollbacka | Visok | **Dizajnersko ograničenje** — klasa vraća samo ListObject podatke (:102-142); inherentno svakom snapshot TX modelu | P2 | Audit *_TX wrappera da su side effecti strogo post-commit; zapisati pravilo | M |
| 5 | Journal može sadržati događaj za poništen upis | Visok | **Tačno** — journal je CSV fajl (modJournaling.bas:58-100), rollback ga ne briše; `CheckJournalForRecovery` poredi broj redova (:218) → lažno „Datenverlust" upozorenje, rizik pogrešnog ručnog reimporta | P2 | Pri rollbacku dopisati storno-marker u journal ili obraditi u recovery runbook-u | S |
| 6 | `RestoreTable` ne invalidira modDataAccess cache | Srednji | **Tačno** — nema poziva invalidacije (:102-142); prozor je read-only report pa je preklapanje malo verovatno | P2 | Javni invalidation hook u modDataAccess + poziv iz RestoreTable po tabeli | S |
| 7 | Asinhroni disk save — crash pre autosave gubi commit | Srednji | **Dizajnersko ograničenje** — AR-002a komentar :71-74, svestan trade-off; journal CSV je kompenzacija za inserte | Prihvaćeno | Dokumentovati; journal pokriva insert recovery, autosave ≤60s | S |
| 8 | Nema `Class_Terminate` zaštite | Srednji | **Tačno** — klasa nema destruktor (ceo fajl); izgubljena referenca sa aktivnim TX ostavlja Excel u manual/no-events | P2 | `Class_Terminate` koji zove `CleanUp` ako je `mActive` | S |
| 9 | `AddTableSnapshot` bez `lo Is Nothing` provere | Srednji | **Tačno** — :44-48 direktno `lo.DataBodyRange` → runtime err 91 za pogrešan naziv | P2 | Guard + `Err.Raise` sa nazivom tabele (jasnija poruka) | S |
| 10 | Redosled rollbacka nije deklarisan | Srednji | **Delimično** — tačno, ali Dictionary čuva insertion order, a value-restore po tabelama je nezavisan (Excel nema FK) → bez praktične posledice | P3 | Ništa; jedna rečenica u komentaru klase | S |
| 11 | Rollback odbija restore pri promeni broja kolona | Nizak | **Dizajnersko ograničenje** — namerna sanity provera (:110-112) | Prihvaćeno | Dokumentovati: schema izmene ne idu kroz ovaj TX | S |
| 12 | Drugi snapshot iste tabele tiho ignorisan | Nizak | **Dizajnersko ograničenje** — :55-57 čuva originalno stanje; savepoint nije cilj klase | Prihvaćeno | Ništa | S |

Bilans: 6 Tačno / 2 Delimično / 0 Netačno / 4 Dizajnersko ograničenje.

### FM-0004 — `modSetup.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | Zeleni setup ne pokriva parcele/palete/korisnike/audit… | Visok | **Delimično** — lista :649-654 stvarno pokriva samo core (18 tabela), ali izostavljeni domeni su opcioni sa svojim Ensure* entry pointovima; „Visok" preteran | P2 | Advisory sekcija u RunSetupHealthCheck: koji opcioni domeni jesu/nisu instalirani | M |
| 2 | `IsSetupHealthy` uži od health-checka, ime zavodi | Visok | **Delimično** — sadržaj tačan (:165-188: samo flag+folderi+tabele), ali funkcija NEMA nijednog pozivaoca u src-vba → rizik pogrešnog zaključka trenutno ne postoji | P3 | Obrisati ili označiti kao rezervisan API uz komentar šta pokriva | S |
| 3 | `EnsureRuntimeSchema` tiho guta greške | Visok | **Tačno** — globalni `On Error Resume Next` :1127; failure stavke bez loga (LogSetup samo na uspeh, :1541) | P2 | Posle svake stavke `Err.Number` provera → `LogSetup "WARN"` + `Err.Clear` | S |
| 4 | Nema schema version ledger-a | Visok | **Dizajnersko ograničenje** — tačno da ne postoji; idempotentni aditivni Ensure* model je svesna, primerena alternativa (FM to priznaje u 5.20) | P2 | `SCHEMA_VERSION` ključ u tblLocalConfig + log primenjenih Ensure* koraka | M |
| 5 | Prvi admin nije transakcioni | Visok | **Tačno** — :1322 AppendRow praznog niza + 7× UpdateCell :1326-1338 bez TX; povratne vrednosti ignorisane; delimičan korisnik potom blokira rerun (dup-check :1304-1310) | P2 | `clsTransaction` oko toka (snapshot TBL_KORISNICI) + provera svake UpdateCell povratne vrednosti | S |
| 6 | Journal prvog admina = prazan niz, UpdateCell nejournalizovan | Visok | **Dizajnersko ograničenje** — činjenično tačno (modDataAccess.bas:210), ali prazan-red+upis-po-imenu je svesni drift-safe obrazac po CLAUDE.md (komentar :1325); jednokratan bootstrap → „Visok" preteran | P3 | Ništa posebno; rešava se eventualnim UpdateCell journal hookom (FM-0002 #2) | S |
| 7 | `tblConfig` legacy, a i dalje obavezna tabela | Srednji | **Tačno** — komentar :10-15 „legacy, validira ako postoji", a :650 je zahteva; isto modMain.bas:286; čista instalacija bez nje pada | P2 | Ukloniti TBL_CONFIG iz oba required spiska (modSetup.bas:650, modMain.bas:286) | S |
| 8 | VeryHidden nije bezbednosna granica | Srednji | **Dizajnersko ograničenje** — komentar :79-81 eksplicitno „anti-tamper" sa dokumentovanim izlazom (ShowConfigSheet); ne tvrdi se security | Prihvaćeno | Dokumentovati da tajne nisu zaštićene od korisnika sa VBA pristupom | S |
| 9 | `BackfillColumn` silent best-effort | Srednji | **Tačno** — :1230 `On Error Resume Next`, bez izveštaja; vrednosti su bezopasni defaulti („Ne", „0", „Aktivan"; prazno ionako = aktivno) | P3 | Brojač neuspeha + `LogSetup "WARN"` | S |
| 10 | Nema jednog entry pointa za punu šemu | Srednji | **Tačno** — SetupNewPC ne zove EnsurePaletniListSchema/Dorade/Korisnici/Audit; redosled je na operateru | P2 | Javni `EnsureFullSchema` koji idempotentno poziva sve Ensure* redom | S |
| 11 | Setup dominantno Windows-specifičan | Srednji | **Dizajnersko ograničenje** — backslash putanje, `Environ$`, pdftotext.exe, Drive for Desktop; Windows Excel je ciljna platforma | Prihvaćeno | Ništa | S |
| 12 | `SetLocalConfigValue` zaobilazi Data Access | Srednji | **Dizajnersko ograničenje** — :456-473 direktan upis; namerno odvojen per-mašina sloj koji mora raditi i pre core šeme | Prihvaćeno | Ništa; audit nepotreban za workstation putanje | S |
| 13 | Additive-only: pogrešno ime kolone → nova kolona | Srednji | **Dizajnersko ograničenje** — :1528-1543 samo dodaje; rename/migracija se rešava namenskim koracima (postojeća praksa u Dorade) | Prihvaćeno | Rename slučajeve i dalje kroz posebne migracione korake; zabeležiti pravilo | S |
| 14 | Setup log bez retention politike | Nizak | **Tačno** — SETUP_LOG :1729-1758, samo append | P3 | Opciono odsecanje na poslednjih N redova u InitSetupLog | S |
| 15 | `EnsureFolder` soft-fail, pozivalac mora proveriti | Nizak | **Tačno** — :1655-1656 loguje bez propagacije; banka folderi se post-proveravaju (:550-552), ali `EnsureAppFolders` (:506-529) nikad ne puni `msg` → setup zelen i bez app foldera (jače od FM tvrdnje) | P2 | `Dir$` post-provera u EnsureAppFolders → u setup report | S |

Bilans: 7 Tačno / 2 Delimično / 0 Netačno / 6 Dizajnersko ograničenje.

### FM-0005 — `modHelpers.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | `filterZbirneKeys` parametar se ignoriše | Visok | **Delimično** — parametar zaista mrtav (:213, telo :216-276 ga ne koristi), ALI jedini pozivalac je modIzvestaj.bas:1523 bez argumenta → nema pogrešnog izveštaja danas; API dug, ne bug | P3 | Ukloniti parametar (ili implementirati filter ako je planiran) | S |
| 2 | Exclude cache: jednak broj redova ≠ puna tabela | Visok | **Tačno** — canCache poredi samo `UBound` (:156-158); sortiran/izveden niz iste dužine može dobiti/upisati tuđ keš; prozor je read-only pa je verovatnoća mala, mehanizam potvrđen | P2 | Keširanje premestiti u novu `GetTableDataNoStorno` (kešira isključivo sopstveni GetTableData izlaz) | M |
| 3 | Funkcije ne proveravaju column indekse pre pristupa | Srednji | **Tačno** — GetVozacDisplayList :41-53, FillComboKooperantiByStanica `colStanica` :117, CheckVerwaisteDokumente svi `col*` | P2 | `If col* = 0 Then Exit` guardovi u tri funkcije | S |
| 4 | Dijagnostika koristi `#,##0` umesto `FmtKolicina` | Srednji | **Tačno** — „izvor istine" :5-11, a `#,##0` na :300, :329, :349, :362, :405, :428 (gube se decimale) | P3 | Zameniti `Format$ "#,##0"` sa `FmtKolicina` u dijagnostici | S |
| 5 | Prikaz 5 stavki, `"..."` tek preko 40 | Srednji | **Tačno** — verwOtp/verwPrij: prikaz do 5 (:347, :360), `"..."` tek `>40` (:351, :364); lostBlk ispravan (:426/:432); ukupan broj ipak stoji u naslovu → blaže | P3 | Promeniti `> 40` u `> 5` na linijama 351 i 364 | S |
| 6 | Utility + UI + document-chain monitoring u istom modulu | Srednji | **Tačno** — potvrđeno uvidom u ceo fajl; poznat arhitektonski dug | Prihvaćeno | Ne cepati sada (minimal-delta pravilo); novu dijagnostiku dodavati u modDokumenta | S |
| 7 | `nz` ne obrađuje Error variant | Srednji | **Tačno** — :193-199 bez `IsError` grane (CStr na CVErr puca); `NzToText` :206 bezbedan; tabele bez formula → retko | P3 | Dodati `IsError` granu u `nz` (vrati default) | S |
| 8 | `BuildManjakDict` može kreirati ključ `""` | Srednji | **Tačno** — :239-240 prazan `brZbr` postaje ključ; prijemnice sa praznim brojem sabiraju se pod `""` (:267-271) | P3 | Preskočiti redove sa praznim brojem zbirne u obe petlje | S |
| 9 | Insertion sort O(n²) za kooperante | Nizak | **Tačno** — :126-136; za realne veličine šifarnika prihvatljivo | Prihvaćeno | Ništa; po rastu preći na postojeći `modArrayUtils.SortArray` | S |
| 10 | Hardkodovani nazivi kolona (vozači, kooperanti) | Nizak | **Tačno** — literali :42-45 i :99-103; `COL_KOOP_ID` postoji u modConfig:107 (neiskorišćen), `COL_VOZ_*` ne postoji | P3 | Koristiti `COL_KOOP_ID`; uvesti `COL_VOZ_*` konstante u modConfig | S |

Bilans: 9 Tačno / 1 Delimično / 0 Netačno / 0 Dizajnersko ograničenje.

**Bilans bloka A (47):** 31 Tačno / 6 Delimično / 0 Netačno / 10 Dizajnersko ograničenje. Nijedna stavka opovrgnuta; preterivanja su u težini (Visok za mrtve/jednokratne puteve). Najhitnije u bloku: `RollbackTx` cleanup (P1/S), fantomski red u `AppendRow` (P1/S), audit snapshot listi po `*_TX` wrapperima (P1/M).

---

## Blok B — Otkup domen (FM-0006…FM-0010, 78 stavki)

### FM-0006 — `modOtkup.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | Pozicioni upis osnovnog reda; reorder tiho korumpira | Kritičan | **Tačno** — `RequireColumns` proverava samo prisustvo (:510-532), `Array`→`AppendRow` poziciono (:550-576) | P2 (latentno; okida tek schema-reorder/migracija, tada P0-klasa) | Upis po imenu preko `GetColumnIndex` mape (obrazac iz `modKooperant.CreateKooperantByName`) | M |
| 2 | Ignorisani `UpdateCell` — izdata amb./vreme/bruto nestaju bez greške | Visok | **Tačno** — :584-592, povratna vrednost se ignoriše; te 3 kolone nisu u `RequireColumns` :510-532 | P2 | Zameni sa postojećim `RequireUpdateCell` ili dodaj 3 kolone u `RequireColumns` | S |
| 3 | `SaveOtkup_TX` bez `tblNovac`/avansa — različita finansijska semantika | Visok | **Delimično** — činjenice tačne (:26-27, nema `SaveNovac`), ali jedini pozivaoci su testovi (`modBusinessFlowProTests:685,710,994,1001`); produkcija koristi samo Multi | P2 | Označi `SaveOtkup_TX` kao test-only ili premesti u test modul | S |
| 4 | Rollback failure se guta | Visok | **Tačno** — `RollbackTx` pod `On Error Resume Next` (:69-96, :330-357); vraća `""` bez signala | P2 | Uhvati grešku rollback-a → `Monitor_Event` ROLLBACK_FAIL + MsgBox | S |
| 5 | Core ne štiti broj dokumenta (prazan/dupli) | Visok | **Delimično** — tačno za modul, ali jedini UI put ima obavezu + `CheckDuplicate` (`frmOtkup:1017-1032`) → težina preterana | P2 | Pre `BeginTx` u Multi pozovi `CheckDuplicate` za neprazan `brDok` | S |
| 6 | Avans snapshot pokrivenost nepotvrđena | Visok | **Netačno** (kao rizik) — provereno: `ApplyAvansToOtkup` piše samo `tblNovac` (`modNovac:1139,1151,1154`) + `tblOtkup` status (:1180); oba snapshotovana (:186-188) | Prihvaćeno | Ništa; u FM upisati da je pokrivenost potvrđena | S |
| 7 | Prosek gajbice izvan TX API-ja | Srednji | **Dizajnersko ograničenje** — funkcija sadrži MsgBox dijaloge (:401-414), ne može u TX; pozivalac dokumentovan (:371) | Prihvaćeno | Opciono: tihi (bez-dijaloga) mod za ne-UI pozivaoce | S |
| 8 | Fail-open kontrola proseka | Srednji | **Dizajnersko ograničenje** — eksplicitno „Fail-safe" (:420-423) | Prihvaćeno | Dodaj Monitor event uz `LogErr` (vidljivost fail-open puta) | S |
| 9 | Nema referential validacije (kooperant/stanica/vozač/parcela) | Srednji | **Tačno** — :458-506 samo ne-prazno, bez lookup-a postojanja | P2 | Opcioni existence-check za kooperanta i stanicu (najbitniji FK) | M |
| 10 | Cena se prihvata od pozivaoca, bez cenovnika | Srednji | **Dizajnersko ograničenje** — cenovnik je predlog (`frmOtkup:396-399`) | Prihvaćeno | — | S |
| 11 | `KulturaID` lookup samo po vrsti | Srednji | **Tačno** — :542-543; `tblKulture` ima red po (vrsta,sorta) → prvi match može biti pogrešna sorta | P2 | Lookup po `VrstaVoca`+`SortaVoca`, fallback postojeći | S |
| 12 | Jednostrani datum filter se tiho ignoriše | Srednji | **Tačno** — :689, :732 traže oba datuma | P2 | Primeni `>=`/`<=` i kad je zadat samo jedan datum | S |
| 13 | `GetSaldoByStation` naziv širi od semantike | Srednji | **Tačno** — bruto zbir; TODO :795-798 to i priznaje | P3 | Preimenovanje ili doc-komentar uz javni API | S |
| 14 | Monitoring hardkoduje `userId="Operator"` | Srednji | **Tačno** — :49, :88, :311, :349 | P3 | `userId` iz `Application.UserName`/config | S |
| 15 | Mešoviti jezici poruka (nemački) | Nizak | **Tačno** — „fehlgeschlagen" :36, :214, :252, :283, :579 | P3 | Prevesti kroz `modPoruke` katalog pri sledećem dodiru | S |
| 16 | Return `"ID1 + ID2"` nije strukturiran | Nizak | **Tačno** — :293-299; ali ustaljen ugovor (parsiraju `modAutoHladnjaca:138`, `modOtkupBlok:1445`) | P3 | Zasad dokumentovati ugovor; dugoročno ByRef out parametri | M |

Bilans: 16/16 — 10 Tačno, 2 Delimično, 1 Netačno (avans pokrivenost zatvorena), 3 Dizajnersko; bez P0/P1, najvažniji hardening = #1 pozicioni insert (P2/M).

### FM-0007 — `frmOtkup.frm`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | Auto-kooperant van otkupnog TX-a; ostaje posle odbijanja | Kritičan | **Delimično** — mehanizam tačan (:1004 pre check-ova :1023/:1064/:1067/:1071; Multi ne snapshotuje `tblKooperanti`), ali posledica je benigni orphan koji retry ponovo pronađe imenom — nije korupcija | P2 | Premesti `ResolveKooperantByName` neposredno pre `SaveOtkupMulti_TX` (posle duplicate/prosek) | S |
| 2 | Pogrešno poređenje kulture parcele | Visok | **Tačno** — :1040-1048 poredi `tblParcele.Kultura` (sorta-semantika, up. :654-659) sa `cmbVrstaVoca` → lažno upozorenje | P1 | Poredi sa `cmbSortaVoca`; fallback lookup sorta→vrsta pa poredi vrstu | S |
| 3 | Neuspešan date re-lock nije blokirajući | Visok | **Tačno** — :711-717: na `False` nema poruke, datum se ne vraća; `btnUnos` ne proverava lock | P1 | Na `False`: MsgBox + vrati `txtDatum` na `GetActiveDatum()` | S |
| 4 | Auto-hladnjača exception tiho progutan | Visok | **Delimično** — OERN omotač postoji (:1128-1135), ali callee ima svoj EH koji uvek vraća warning (`modAutoHladnjaca:197-201`) → prozor vrlo uzak | P2 | Posle poziva proveri `Err.Number` pre `On Error GoTo 0`, prikaži generičko upozorenje | S |
| 5 | Post-save linking bloka best-effort, tiho | Visok | **Tačno** — :1174-1176 OERN; `AfterUnos` EH samo LogErr (`modOtkupBlok:271-273`); nevezan blok NE ulazi u „Izgubljene" (`modDokumenta:2856` preskače prazan link) | P1 | `AfterUnos`/link vrate Boolean → MsgBox „blok NIJE vezan"; „Izgubljeni" da uključi prazan link | S |
| 6 | Storno dugme bez akcije | Visok | **Tačno** — :1228-1230 samo `ButtonActive` = čisto stilizovanje (`modTheme:342-345`) → mrtvo dugme | P2 | Sakrij/disable dugme ili poveži na panelski storno tok | S |
| 7 | Parcela nije hard-validirana (override; bez aktivnost/GGAP) | Srednji | **Tačno** — vbYesNo override :1048-1055; lista bez filtera :623-634 | P2 | Uz fix #2; opcioni filter neaktivnih ako kolona postoji | S |
| 8 | Auto-cena nije enforced | Srednji | **Dizajnersko ograničenje** — „rucni unos ostaje moguc" (:396-397) | Prihvaćeno | — | S |
| 9 | Primalac nije obavezan uz gotovinu | Srednji | **Tačno** — :984-991 samo numerika | P2 | Zahtevaj `txtPrimalac` kad je `novac>0` (uz `IsKesIsplate`) | S |
| 10 | Duplicate check nije unique constraint (race) | Srednji | **Delimično** — tačno (:1023-1032 odvojen read), ali single-writer + station lock → mala verovatnoća | P2 | Jeftino: ponovi `CheckDuplicate` neposredno pre TX poziva | S |
| 11 | Dynamic-control failure fail-open | Srednji | **Tačno** — EH :195-198 LogErr + `Nothing`; polje tiho nestaje (guard :877) | P2 | Jednokratna poruka operateru kad runtime kontrola ne uspe | S |
| 12 | Stari broj zbirne se zadržava | Srednji | **Delimično** — namerno za serijski unos; reset tačke postoje (:523-531, :749-755); rizik samo ručno uneta zbirna | P2 | U `ClearOtkupFields` očisti `txtBrojZbirne` kad panel kontekst nije aktivan | S |
| 13 | Strict validation može biti isključena | Srednji | **Dizajnersko ograničenje** — `VALIDACIJA_UNOSA` toggle (:795, :908-934, :1017); bruto režim i dalje traži gajbe (:922-933) | Prihvaćeno | Dokumentuj u uputstvu šta OFF tačno isključuje | S |
| 14 | Auto-chain nije atomski sa otkupom | Srednji | **Dizajnersko ograničenje** — namerni post-commit best-effort + recovery (:1102-1114) | Prihvaćeno | — | S |
| 15 | UserForm ima previše odgovornosti | Nizak | **Tačno** — 1294 linija orchestration-a | P3 | Prihvatiti; bez refaktora (minimal-delta politika) | L |
| 16 | Mešoviti jezici/encoding komentara | Nizak | **Tačno** — nemački komentari (:19-23, :378-387) | P3 | Postepeno prevoditi pri dodiru koda | S |

Bilans: 16/16 — 9 Tačno, 4 Delimično, 3 Dizajnersko; tri P1 (parcelno poređenje #2, tihi date-lock #3, tihi link bloka #5), sve S napor.

### FM-0008 — `modKooperant.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | Identitet samo po imenu; iste osobe se spajaju | Kritičan | **Tačno** — :57-82 samo Ime+Prezime, prvi match, bez BPG/JMBG/disambiguacije | P1 | Kod >1 pogotka blokiraj free-text („izaberite iz liste"); kasnije disambiguation dijalog | M |
| 2 | Name lookup vraća i neaktivnog | Visok | **Tačno** — :62-79 ne čita `Aktivan`; free-text zaobilazi UI filter | P2 | U petlji preskoči `Aktivan<>STATUS_AKTIVAN` | S |
| 3 | Auto-create i otkup nisu atomski | Visok | **Delimično** — isto kao FM-0007 #1 (dva TX-a: :116-125 pa Multi; orphan benigni) | P2 | Isto kao FM-0007 #1 (reorder u formi) | S |
| 4 | Cross-station fallback zaobilazi filter | Visok | **Delimično** — kod tačan (:76-81), ali ponašanje dokumentovano namerno (:55-56); neusklađenost sa `KOOP_FILTER_BY_OM` stoji | P2 | Uz uključen filter traži potvrdu operatera pre cross-station rezultata | S |
| 5 | Concurrent `GetNextID` bez lock/retry | Visok | **Delimično** — mehanizam tačan (:93), ali single-writer desktop → nizak realan rizik | P2 | Ništa sada; oslonac na postojeći lock/sync model | M |
| 6 | Duplicate race (dva korisnika, isto ime) | Srednji | **Delimično** — ista multi-user klasa kao #5 | P2 | — (isto kao #5) | M |
| 7 | Slaba normalizacija imena (razmaci/Unicode) | Srednji | **Tačno** — :67-71 samo `LCase`+`Trim`; dupli unutrašnji razmak pravi duplikat | P2 | Kolabiraj višestruke razmake pre poređenja i pre kreiranja | S |
| 8 | Nedostajuća kolona → generička `row(0)` greška | Srednji | **Tačno** — :110-114 indeksi bez provere; pukne pre `BeginTx` (:116), poruka generička (:133) | P2 | `RequireColumns` za 5 kolona pre mapiranja | S |
| 9 | Eventual sync novog partnera | Srednji | **Dizajnersko ograničenje** — eksplicitno dokumentovano (:12-13) | Prihvaćeno | — | S |
| 10 | Bound ID se ne revalidira | Srednji | **Delimično** — tačno (:28-32), ali lista se gradi iz iste tabele i redovi se ne brišu → zastareo ID praktično nemoguć | P3 | Ništa (aktivnost pokriva fix #2) | S |
| 11 | Rollback failure maskira originalnu grešku | Srednji | **Tačno** — EH :130-133: `RollbackTx` bez lokalnog OERN → preskače LogErr/MsgBox | P2 | `On Error Resume Next` oko `RollbackTx` (obrazac iz `modOtkup`) | S |
| 12 | Minimalni matični podaci bez statusa „nepotpun" | Nizak | **Dizajnersko ograničenje** — namerni minimalni onboarding (:10-13) | Prihvaćeno | Opciono status „nepotpun" za naknadnu dopunu | S |
| 13 | `SplitName` heuristika | Nizak | **Tačno** — :139-149 prvi razmak; svesno i dokumentovano | P3 | — | S |

Bilans: 13/13 — 6 Tačno, 5 Delimično, 2 Dizajnersko; jedini P1 = #1 (identitet po imenu — poslovni rizik pogrešnog knjiženja).

### FM-0009 — `modOtkupBlok.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | Link ne proverava domain kompatibilnost | Kritičan | **Delimično** — guard-ovi samo postojanje/storno/prazan link (:1441-1461); ali stanica-mismatch već pokriven UI resetom (`frmOtkup:523-531`) → ostaje datum/proizvod; „Kritičan" preterano | P2 | U link TX uporedi vrstu (i datum) otkupa vs otpremnice; razlika → preskoči + upozori | M |
| 2 | `OtpremnicaID`/`BrojZbirne` divergencija | Kritičan | **Delimično** — tačno da TX piše samo `OtpremnicaID` (:1465), zbirna iz prefill-a (:722); ali traži ručnu izmenu zbirne posle prefill-a → uzak scenario; recovery menja oba (:1061) | P2 | U link TX upiši i `BrojZbirne` sa ciljne otpremnice kad je prazan/različit | S |
| 3 | Post-save link nije atomski sa otkupom | Visok | **Dizajnersko ograničenje** — svesna post-save orkestracija (zajednički TX zahtevao bi spajanje sa Multi) | P2 | Vidljivost umesto atomicnosti (vidi #4) | M |
| 4 | Link failure samo logovan | Visok | **Tačno** — EH :1475-1477 rollback+LogErr; `AfterUnos` :271-273 isto; nevezan blok ne ulazi u „Izgubljene" (`modDokumenta:2856`) → rizik duplog unosa | P1 | Link vrati Boolean → MsgBox „blok NIJE vezan" + uputstvo za recovery | S |
| 5 | Stornirani blok daje default cenu | Visok | **Tačno** — `ExistingBlokCena` :1600-1610 i `BuildFirstBlokCena` :1613-1628 bez storno filtera (up. :1637 koji filtrira) | P1 | `ExcludeStornirano`/storno-check u oba helpera | S |
| 6 | Correction prefill po samom `BrDok` | Visok | **Tačno** — :958-967 prvi I/II red bez storno/generacije; `CheckDuplicate` preskače stornirane (`modDataAccess:414`) pa je reuse broja realan | P2 | Uzmi poslednju generaciju (poslednji red) i/ili filtriraj na upravo stornirane | S |
| 7 | Overfill dozvoljen override-om | Visok | **Dizajnersko ograničenje** — warning sa vbYesNo (:227-232), namerna semantika | Prihvaćeno | Opciono: Monitor event pri override-u (audit trag) | S |
| 8 | Confirm fail-open | Visok | **Dizajnersko ograničenje** — dokumentovano „nikad ne blokira unos" (:178-179, :234-236) | P2 | Monitor event kad kontrola pukne (vidljivost) | S |
| 9 | Partial prefill tih (OERN) | Srednji | **Tačno** — ceo `PrefillLeftForm` pod OERN (:713-741), bez završne validacije | P2 | Posle prefill-a proveri datum/OM; neuspeh → poništi izbor + poruka | S |
| 10 | Malina „Kupac" = prvi kooperant | Srednji | **Dizajnersko ograničenje** — pretpostavka „1 koop. po otpremnici" dokumentovana (:1656-1657) | P3 | Prikaži „+N" kad ima više kooperanata | S |
| 11 | Spec. po datumu otkupa, ne otpremnice | Srednji | **Delimično** — činjenično tačno (:1269-1273), ali za specifikaciju blokova prirodan izbor; očekivanje korisnika spekulativno | P3 | Naznači „po datumu otkupa" u subtitle | S |
| 12 | Session cena nije persistirana | Srednji | **Dizajnersko ograničenje** — svestan session default (:48, :2127); posle prvog bloka cena živi u podacima | Prihvaćeno | — | S |
| 13 | Schema guards neujednačeni | Srednji | **Tačno** — `LoadOtpremnice`/`LoadBlokovi` bez 0-provere (:501-507, :619-626) vs :1669 koji proverava | P2 | `RequireColumnIndex` u load funkcijama panela | S |
| 14 | Višestruki O(n) scanovi | Srednji | **Tačno** — tri zasebna prolaza `SumKol/Bruto/AmbByOtp` (:1542-1598) po refresh-u; request-cache ublažava | P3 | Jedan kombinovani prolaz za tri sume | S |
| 15 | `mPrefilling` lifecycle | Nizak | **Tačno** — `Release` :2122-2161 i `Attach` :103-116 ga ne resetuju; zaglavljivanje praktično nemoguće (OERN tok uvek stigne do reseta) | P3 | `mPrefilling=False` u Attach i Release (1 linija) | S |
| 16 | Module-level singleton state | Nizak | **Dizajnersko ograničenje** — jedna `frmOtkup` instanca po sekciji (:44-45) | Prihvaćeno | — | S |

Bilans: 16/16 — 7 Tačno, 3 Delimično (uklj. oba „Kritična" ublažena na P2), 6 Dizajnersko; dva P1 (tihi link failure #4, storno-cena default #5), oba S napor.

### FM-0010 — `modAutoHladnjaca.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | `outBrPrij` = generisan, ne kreiran broj | Kritičan | **Tačno** — :150-151 postavljen pre svakog `SavePrijemnica_TX`; pri padu ostaje popunjen → `frmOtkup:1143` guard nikad ne okida, relink ide na nepostojeći broj (Reassign padne uz zbunjujuću poruku) | P1 | Postavi `outBrPrij` tek posle uspešnog `SavePrijemnica_TX` | S |
| 2 | Backfill razdvaja brojeve klasa iste zbirne | Kritičan | **Tačno** — `brByZbr` :343-360 puni se samo tokom run-a; postojeća prijemnica druge klase se ne konsultuje → novi broj krši konvenciju | P1 | Pre petlje napuni `brByZbr` brojevima postojećih prijemnica po zbirnoj | S |
| 3 | Nema end-to-end atomiciteta (do 6 TX) | Visok | **Dizajnersko ograničenje** — namerna saga + warning + repair alati (:154-183) | Prihvaćeno | — | L |
| 4 | Link failure nije propagiran | Visok | **Tačno** — EH :240-243 rollback+LogErr, bez rezultata; lanac javlja „kompletan" | P1 | Link vrati Boolean; pad → dodaj u warning tekst lanca | S |
| 5 | Otpremnica rezultat nevalidiran | Visok | **Tačno** — `otpID` (:156, :172) se ne proverava; prazan → zbirna+prijemnica ipak nastaju, link piše samo zbirnu (:231) | P1 | `Len(otpID)=0` → `failKlase` + preskoči ostatak klase | S |
| 6 | Zbirna rezultat nevalidiran | Visok | **Tačno** — `SaveZbirna_TX` kao statement, rezultat odbačen (:158-159, :174-175) | P1 | Proveri rezultat; prazan → warning kao za prijemnicu | S |
| 7 | Mirror vozač best-effort | Visok | **Tačno** — :112-119 OERN oko `Ensure...`; `vozacID=stanicaID` i kad mirror nije kreiran | P2 | Posle Ensure proveri postojanje vozača pre nastavka | S |
| 8 | Backfill kriterijum preširok | Visok | **Tačno** — :319-328 bez provere kupca zbirne (interni tok) → prijemnica može nastati za eksterni tok | P2 | Kandidat samo ako `zbirna.KupacID == MALINA_DEFAULT_KUPAC` | S |
| 9 | Dupli kandidati u jednom run-u | Visok | **Tačno** — `have` se ne dopunjava posle uspeha (:370) | P2 | Posle uspešnog save-a `have(key)=True` (1 linija) | S |
| 10 | Config kupac se ne validira | Srednji | **Tačno** — :96-107 samo ne-prazan; postojanje/aktivnost bez provere | P2 | `LookupValue` postojanja kupca; nema → isto upozorenje kao prazan config | S |
| 11 | Link prepisuje postojeću vezu | Srednji | **Tačno** — :231-232 bezuslovno (čuva samo `VozacID` :233-234); privatan helper, jedno pozivno mesto u ispravnom redosledu | P3 | Opcioni existing-link guard (pažljivo — correction tok) | S |
| 12 | String contract OtkupID-jeva | Srednji | **Tačno** — `Split " + "` + `hasKlasaI` (:135-145); krhko ali stabilan interni ugovor | P3 | Dugoročno ByRef `idI`/`idII` iz Multi_TX | M |
| 13 | Pending relink samo u memoriji | Srednji | **Dizajnersko ograničenje** — svesno in-memory stanje (:34-36), čisti se na Terminate; ručni orphan alat je fallback | Prihvaćeno | — | S |
| 14 | Backfill praznu klasu pretvara u I | Srednji | **Tačno** — `ClassOrDefault` :402-405; legacy-friendly, maskira izvorni podatak | P3 | Loguj kandidate sa praznom klasom | S |
| 15 | Orphan palete traže ručnu pripremu | Srednji | **Tačno** — samo upozorenje (:336-340), bez preflight-a | P3 | Preflight lista konfliktnih brojeva u `tblPaletaStavka` pre potvrde | M |
| 16 | Hladnjača helperi fail-soft | Nizak | **Tačno** — :47-67 OERN → greška izgleda kao poslovno `False` (lanac tiho preskočen) | P3 | Dodaj `LogErr` u oba helpera | S |
| 17 | Fallback broj sa sekundama | Nizak | **Delimično** — :126 `hhnnss`; kolizija traži dva unosa iste sekunde bez `brDok` — praktično nemoguće single-writer | P3 | Ništa (ili sufiks OtkupID) | S |

Bilans: 17/17 — 14 Tačno, 1 Delimično, 2 Dizajnersko; najrizičniji modul: pet P1 (#1 `outBrPrij`, #2 backfill brojevi, #4 tihi link, #5/#6 nevalidirani rezultati), svi popravljivi S naporom.

**Bilans bloka B (78):** 46 Tačno / 15 Delimično / 1 Netačno / 16 Dizajnersko-Prihvaćeno. Bez P0; 11 P1 — najisplativiji paket: `modAutoHladnjaca` result/link validacije (4×S), storno-cena default u panelu blokova, parcelno poređenje kulture u `frmOtkup`.

---

## Blok C — Dokumenta i storno stack (FM-0011…FM-0015, 91 stavka)

**Ključni cross-check nalazi:** `tblPaletaIstorija` ne postoji nigde u kodu (globalni grep prazan); `Monitor_Event` je interno pod `On Error Resume Next` (modMonitoring.bas:66); `LookupValue` vraća **prvi** match (modDataAccess.bas:459-463); auto-brojevi dokumenata sadrže stanicu+datum (`FormatBroj`, modBrojevi.bas); frmDokumenta za otkup/otpremnicu/prijemnicu zove by-broj wrappere (frm:3193/3203/3218), a za novac šalje **broj** u `StornoNovac_TX` koji očekuje NovacID (frm:3230 vs modStorno.bas:765).

### FM-0011 — `modDokumenta.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | Auto-chain ne proverava upstream TX rezultate | Kritičan | **Tačno** — modAutoHladnjaca.bas:156-158,172-174: `otpID` neproveren, `SaveZbirna_TX` statement bez rezultata; wrapperi vraćaju `""` (modDokumenta.bas:222-224,542-546); samo prijemnica proverena | P1 | modAutoHladnjaca: proveriti otpID/zbirna rezultat, prekinuti lanac + proširiti failKlase upozorenje | S |
| 2 | Positional insert otpremnice/zbirne | Kritičan | **Tačno** — modDokumenta.bas:252-254 i 614-619: `Array(...)`+`AppendRow` bez RequireColumns; prijemnica ima builder (1154+) | P2 | Preći na column-mapped builder po uzoru na `BuildPrijemnicaRowData` | M |
| 3 | Prosek gajbe uključuje storno | Visok | **Tačno** — `SumByBroj` modDokumenta.bas:1623-1641: nema storno filtera; pozivi 1575-1577, 1605-1607 | P1 | U `SumByBroj` preskočiti redove `Stornirano="Da"` (jedan uslov u petlji) | S |
| 4 | Reassign nema domain validator | Visok | **Tačno** — modDokumenta.bas:3151-3157, 3202-3209: samo postoji+nije storniran; bez stanica/datum/vrsta/klasa poređenja | P2 | Zajednički compat-check (stanica, vrsta, datum) sa eksplicitnim override parametrom | M |
| 5 | Recovery liste prikazuju samo prvu class količinu | Visok | **Tačno** — `GetAktivneZbirne` modDokumenta.bas:2984-2991 i `GetOsirocenePrijemnice` 2934-2941: `seen`+prvi red, bez agregacije | P2 | Agregirati količinu preko class redova, kao u `GetAktivnePrijemnice` | S |
| 6 | Optional bruto failure je tih | Visok | **Tačno** — modDokumenta.bas:260 i 1118: `UpdateCell ... COL_*_BRUTO` bez provere Boolean rezultata | P2 | `If Not UpdateCell(...) Then Err.Raise` — unutar TX-a pa je rollback automatski | S |
| 7 | Prijemnica bez zbirne u warning režimu | Visok | **Dizajnersko ograničenje** — modConfig.bas:903-911: default je BLOK; warning je svesni opt-in (komentar modDokumenta.bas:2237-2241) | Prihvaćeno | Monitor WARN event pri upisu prijemnice bez postojeće zbirne u warning modu | S |
| 8 | Save core nema duplicate zaštitu | Visok | **Tačno** — Validate* (modDokumenta.bas:2129-2270) nemaju unique proveru broja+klase | P2 | Guard u `Save*`: odbij aktivan isti broj+klasa | M |
| 9 | Paleta istorija nije u snapshot listi | Visok | **Netačno** — `tblPaletaIstorija` ne postoji u kodu; `PaletizePrijemnica` piše samo tblPaleta/tblPaletaStavka, obe u snapshotu (modDokumenta.bas:850-851) | Prihvaćeno | Nema akcije — snapshot lista je kompletna | — |
| 10 | Manjak fallback izgleda kao nulti manjak | Srednji | **Tačno** — modDokumenta.bas:1475-1476: EH vraća `Array(0, pending, 0, 0)` | P3 | Na grešku vratiti Empty/flag; UI da prikaže „obračun nije uspeo" | S |
| 11 | Referencijalni ID-jevi se ne proveravaju | Srednji | **Tačno** — validacije samo `Len(Trim$)>0` (2141-2148) | P2 | Existence lookup za stanicu/vozača/kupca u Validate* | S |
| 12 | Vrsta/sorta nisu core-obavezne | Srednji | **Tačno** — Validate* uopšte ne primaju vrsta/sorta | P3 | Poslovna odluka; eventualno WARN log na prazno | S |
| 13 | Cena 0 je dozvoljena | Srednji | **Tačno** — 2161, 2258: samo `cena < 0` pada | P3 | Eksplicitno dokumentovati pravilo ili config flag | S |
| 14 | Otvoreni datum interval nije podržan | Srednji | **Tačno** — modDokumenta.bas:376: `datumOd > 0 And datumDo > 0` | P3 | Podržati jednostrani opseg u filter grani | S |
| 15 | Historical chain može spojiti generacije | Srednji | **Tačno** — `BuildChainIndex` 2724-2785: bez storno filtera, ključ=poslovni broj; verovatno namerno | P3 | Označiti stornirane u prikazu ili dokumentovati nameru | S |
| 16 | Source active status se ne proverava u reassign-u | Srednji | **Tačno** — 3164-3171: `FindRows` po OtkupID bez storno provere izvora | P2 | Preskočiti/odbiti stornirane izvorne redove u `ReassignOtkupToOtpremnica_TX` | S |
| 17 | Modul ima previše odgovornosti | Nizak | **Tačno** — 3267 linija | P3 | Bez refaktora sada; postepeno izdvajanje read-modela | L |
| 18 | Mešoviti jezici poruka | Nizak | **Tačno** — „fehlgeschlagen fuer tblZbirna" + srpski | P3 | Postepeno kroz modPoruke katalog | S |

Bilans: 18 — 16 Tačno, 1 Dizajnersko, 1 Netačno (#9). Hitnost: 2×P1 (#1 auto-chain, #3 prosek), 7×P2, 7×P3, 2×Prihvaćeno.

### FM-0012 — `modStorno.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | Single-row API na višeklasnom dokumentu | Kritičan | **Delimično** — dualni API postoji (35/69, 190/240, 509/597), svi UI pozivaoci koriste by-broj (frmDokumenta.frm:3193/3203/3218; modOtkupBlok.bas:858-860); realan srodni bug: frm:3230 šalje broj u `StornoNovac_TX` → storno novca uvek pada | P2 (novac fix P1) | Guard u single-ID storno: druge aktivne redove istog broja → greška; popraviti frm:3230 | S |
| 2 | Storno po samom poslovnom broju | Kritičan | **Delimično** — scan bez stanica/datum scope-a (103-114), ali auto-broj sadrži stanicu+datum pa je kolizija realna samo za ručne/eksterne brojeve | P2 | Match-evi sa >1 stanicom → blokiraj i traži uži kontekst | S |
| 3 | Hladnjača gate iz poslednjeg match-a | Kritičan | **Delimično** — mehanika tačna (107 prepisivanje, 124-125 gate), preduslov je kolizija broja iz #2 | P2 | Prekid sa greškom ako match-evi pripadaju različitim stanicama | S |
| 4 | Nema razloga storna | Visok | **Tačno** — `MarkRowStornirano` 1368-1372 samo flag; nijedan API ne prima razlog | P2 | Opcioni reason parametar → monitoring payload + tblStornoVeze | M |
| 5 | Nema potvrđene autorizacije | Visok | **Tačno** — nijedna modAuth/role referenca u fajlu | P2 | Centralni `RequireStornoRight()` hook u *_TX wrappere | M |
| 6 | Novac ostaje aktivan posle storna otkupa/fakture | Visok | **Dizajnersko ograničenje** — `ResetNovac*Link` 1121-1159 samo odvezuje; svesni finansijski model | Prihvaćeno | UI poruka posle storna: „isplata ostaje aktivna — storniraj zasebno ako treba" | S |
| 7 | Inconsistent Fakturisano/FakturaID | Visok | **Tačno** — 576-584: orphan logika samo uz flag „DA"; `fakturaID` pročitan (572) ali neiskorišćen | P2 | Uslov proširiti: `flag="DA" Or Len(fakturaID)>0` | S |
| 8 | Otpremnica Malina kaskada je global-mode based | Visok | **Nije proverivo statički** — kod tačan (250-256, 301-306: `IsMalinaMode()` globalno); zavisi od podataka/konfiguracije | P2 | Kaskadu usloviti i 1:1 proverom | S |
| 9 | Paletne stavke ostaju orphan | Visok | **Dizajnersko ograničenje** — komentar 388-390: „NE dira tblPaletaStavka"; recovery alati postoje | Prihvaćeno | Nema akcije u core | — |
| 10 | Storno audit nije row-journal | Visok | **Tačno** — samo flag + ModifiedAt/By + monitoring | P2 | Append-only storno ledger ili proširiti tblStornoVeze na sve tipove | M |
| 11 | Invoice-level orphan marker se prepisuje | Srednji | **Tačno** — `MarkFakturaOrphaned` 1091-1092: jedan `OsirocenoOd`; stavke čuvaju detalj | P3 | Ne prepisivati popunjen marker | S |
| 12 | `LookupActiveID` vraća poslednji duplikat | Srednji | **Tačno** — 1002-1010: prepisivanje po match-u, bez detekcije | P2 | ByRef `matchCount` ili raise na >1 (hrani FM-0013 defekte) | S |
| 13 | Storno prerade tiho preskače missing paletu | Srednji | **Tačno** — 928-933: Nothing/0 → skip | P3 | Log na missing; raise na count>1 | S |
| 14 | Recovery nije end-to-end rollback | Srednji | **Dizajnersko ograničenje** — orphan/correction model dokumentovan i konzistentan | Prihvaćeno | Nema akcije | — |
| 15 | Paleta istorija ne dobija storno događaj | Srednji | **Netačno** — tabela istorije paleta ne postoji u kodu | Prihvaćeno | Nema akcije | — |
| 16 | Rollback failure se guta | Srednji | **Tačno** — `HandleStornoTxError` 1419-1423: rollback pod OERN, caller vidi samo `False` | P2 | Uhvatiti rollback grešku → poseban Monitor ERROR + tekst u poruci | S |
| 17 | Monitoring user je `Operator` | Srednji | **Tačno** — 1440, 1464 hardkod; `modAuth.GetCurrentUser` postoji | P3 | `userId:=CurrentUser()` helper | S |
| 18 | `IsStorniranoValue` direktan `CStr` | Nizak | **Tačno** — 1383-1385; Error ćelija → greška, EH hvata → rollback (bez korupcije) | P3 | NzTx-stil guard | S |
| 19 | Header pravila su zastarela | Nizak | **Tačno** — 16-18 „nema kaskade" vs 121-151 (hladnjača) i 298-306 (malina) | P3 | Ažurirati header komentar | S |

Bilans: 19 — 11 Tačno, 3 Delimično, 3 Dizajnersko, 1 Netačno (#15), 1 Nije proverivo (#8). Hitnost: 10×P2, 5×P3, 4×Prihvaćeno (novac-identity bug u frm:3230 = P1, evidentiran u FM-0018).

### FM-0013 — `modStornoFlow.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | Multi-class relink na jedan `newOtpID` | Kritičan | **Tačno** — modStornoFlow.bas:462 (`LookupActiveID` = poslednji ID) + 474-477: svi blokovi obe klase → isti `newOtpID` | P1 | U relink petlji čitati klasu bloka i mapirati na odgovarajući novi class OtpremnicaID | M |
| 2 | Context failure ne blokira mutaciju | Kritičan | **Tačno** — neprovereno u 5 grana: otp PONISTENJE 399-402, zbirna DUPLI 607-611, zbirna PONISTENJE 646-649, revers ISPRAVKA 759-762, revers DUPLI/PON 773-776 | P1 | `If Len(cid)=0 Then Exit Function` u 5 grana (obrazac iz ISPRAVKA grana) | S |
| 3 | Paletni detach false-success | Kritičan | **Tačno** — 1209-1216: suma detach rezultata bez provere, `res("ok")=True` bezuslovno; detach vraća 0 i na grešku | P1 | Proveriti svaki detach (0/greška → `MarkCorrectionManual`, ne COMPLETED) | S |
| 4 | Poslovni broj kao root | Kritičan | **Delimično** — `GetOtpremnicaIDsByBroj` 1225-1251 namerno uključuje stornirane generacije; cross-stanica kolizija mala; višestruke ispravke istog broja mogu zahvatiti šire | P2 | Čuvati stare class ID-jeve u correction context i relinkovati samo njih | M |
| 5 | Completion nije atomski | Visok | **Dizajnersko ograničenje** — saga: svaki korak na padu ide u `MarkCorrectionManual` (478-481, 507-510) | Prihvaćeno | Nema akcije; dokumentovan saga ugovor | — |
| 6 | Nova otpremnica nije jednoznačno razrešena | Visok | **Tačno** — 462: bez klase/stanice/datuma, poslednji match | P2 | Razrešiti po klasi+stanici; upozoriti na >1 aktivan match (uz #1) | S |
| 7 | Relink zbirne rezultat se ignoriše | Visok | **Tačno** — 692: statement poziv; helper vraća 0 i na grešku (931-934); recalc potom postavi novu zbirnu na 0/0 pa invariant `0=0` prođe → lažni COMPLETED | P1 | Proveriti vraćeni broj vs `CountActive` stare zbirne; 0 uz očekivano >0 → MANUAL | S |
| 8 | Otkup denorm update je preširok | Visok | **Tačno** — 918-924, 953-968: svi aktivni otkupi po samom starom broju zbirne | P2 | Filtrirati i po `OtpremnicaID ∈ prevezene otpremnice` | S |
| 9 | Nema autorizacije u workflow core-u | Visok | **Tačno** — nijedna modAuth/role referenca | P2 | Centralni auth hook na `Run*`/`Complete*` ulazima | M |
| 10 | Nema operatorovog razloga | Visok | **Tačno** — javni API bez reason parametra | P2 | Opcioni `reason` → `CreateCorrectionContext` message | S |
| 11 | Dvofazno poništenje nije atomsko sa paletama | Visok | **Dizajnersko ograničenje** — Faza A commit 1205, Faza B kroz paletni motor (dokumentovano 1146-1153); problem je isključivo #3 | Prihvaćeno (uz #3 fix) | Nema akcije osim #3 | — |
| 12 | `valid` result contract ne postoji | Srednji | **Tačno** — header :19 vs `NewRes` 1447-1456 (bez `valid`) | P3 | Izbaciti `valid` iz header komentara | S |
| 13 | Scan može spojiti generacije | Srednji | **Tačno** — `ScanOtpremnica` 1304-1318 | P3 | Dokumentovati; već konzistentno u DUPLI | S |
| 14 | `0/False` ne razlikuje prazno od greške | Srednji | **Tačno** — `CountActive` 1412-1414 vraća 0 na grešku; `RecalcOrStornoEmptyZbirna_TX` 1050-1053 na 0 → **storno** zbirne (destruktivan edge) | P2 | `CountActive` raise ili -1 na grešku; caller proverava | S |
| 15 | Ownership funkcija zbunjujuće imenovana | Srednji | **Tačno** — 1014-1021 | P3 | Preimenovati uz stari alias | S |
| 16 | Context odvojen od business mutacije | Srednji | **Dizajnersko ograničenje** — namerno (pending trag mora preživeti rollback) | Prihvaćeno | Nema akcije | — |
| 17 | Format/poruke nisu lokalizovane | Nizak | **Tačno** — ASCII poruke van modPoruke | P3 | Postepeno kroz modPoruke | S |

Bilans: 17 — 13 Tačno, 1 Delimično, 3 Dizajnersko. Hitnost: 4×P1 (#1, #2, #3, #7 — najozbiljniji nalazi celog audita), 7×P2, 4×P3, 3×Prihvaćeno.

### FM-0014 — `modStornoContext.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | Duplicate ID read/write split | Kritičan | **Tačno** — `GetCorrectionRowByID` :206 poslednji match; `GetCorrectionField` :216 → `LookupValue` = **prvi** match; preduslov: dupli COR-ID | P2 | `GetCorrectionField` čitati preko istog reda; raise na count>1 | S |
| 2 | CorrectionID `max+1` concurrency | Kritičan | **Delimično** — mehanika tačna (:44), ali single-writer desktop | P2 | Post-append verifikacija jedinstvenosti ili vremenski sufiks | S |
| 3 | Nema state-transition guard-a | Visok | **Tačno** — Complete/SetTerminalState 100-116, 165-183: bez čitanja trenutnog statusa | P2 | `expectedStatus` parametar; zabraniti terminal→terminal bez override-a | S |
| 4 | Multi-user race | Visok | **Delimično** — nema claim/CAS, ali single-writer | P2 | Pokriveno fixom #3 | S |
| 5 | Positional insert | Visok | **Tačno** — :50-68 `rowData(0 To 17)`; redosled vezan za `EnsureStornoVezeSchemaCore` | P2 | Column-mapped upis za recovery ledger | M |
| 6 | Context nema class mapping | Visok | **Tačno** — jedno `OldDocID`/`NewDocID` polje (:54-59); strukturni uzrok FM-0013 #1 | P2 | CSV lista class-ID parova ili nova kolona | M |
| 7 | Stale recovery text | Visok | **Tačno** — Complete :110-112 ne briše `RecoveryAction`; Cancel šalje "" | P3 | Complete/Cancel eksplicitno prazniti RecoveryAction | S |
| 8 | Resolver korisnik se ne čuva | Visok | **Tačno** — schema tačno 18 kolona bez `CompletedBy`; TBL_STORNO_VEZE nije u `AuditableTables` (modSetup.bas:1082-1089) | P2 | `ResolvedBy` kolona + upis u `SetTerminalState` | S |
| 9 | Latest pending nije stvarno latest | Visok | **Tačno** — `FindLatestPending` :297: poslednji u redosledu tabele, bez `CreatedAt` | P3 | Birati max(`CreatedAt`) | S |
| 10 | Pending lookup preslabo scoped | Visok | **Delimično** — filter tip+mode (:294-296), safe-stop na >1 pending; pogrešan izbor samo uz tačno 1 tuđi pending | P2 | Dodatno filtrirati po `oldBroj` | S |
| 11 | `CompletedAt` semantika | Srednji | **Tačno** — :174: upis i za FAILED/MANUAL/CANCELLED | P3 | Dokumentovati kao `TerminalAt` ili razdvojiti | S |
| 12 | Current-state umesto event history | Srednji | **Dizajnersko ograničenje** — monitoring nosi događaje | Prihvaćeno | Nema akcije | — |
| 13 | Recovery lista može biti parcijalna | Srednji | **Tačno** — :255 direktan `CStr`; `TxCell` :361-363 bez IsError/IsNull; EH → delimična kolekcija | P2 | NzTx-stil guard u `TxCell` i cNeeds ćeliju | S |
| 14 | Badge može lažno pokazati nulu | Srednji | **Tačno** — `CountPendingRecovery` :276-281 pod OERN | P3 | -1/flag na grešku; badge „?" | S |
| 15 | Status/NeedsRecovery nema validator | Srednji | **Tačno** — konzistencija-check ne postoji | P3 | Mali check u health-check | S |
| 16 | Ensure schema je best-effort | Srednji | **Tačno** — :344-350 OERN bez post-check-a | P3 | Post-check `GetTable`+kolone, raise sa porukom | S |
| 17 | Global safe-stop blokira nepovezane korisnike | Srednji | **Dizajnersko ograničenje** — namerno konzervativno (:307-309) | Prihvaćeno | Nema akcije | — |
| 18 | Monitoring nema user attribution | Nizak | **Tačno** — `LogContext` :365-372 bez userId | P3 | `userId:=CurrentUser()` | S |

Bilans: 18 — 13 Tačno, 3 Delimično, 2 Dizajnersko. Hitnost: 0×P1, 10×P2, 6×P3, 2×Prihvaćeno.

### FM-0015 — `modDokumentInvariant.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | Nepostojeća zbirna je validna (`0=0`) | Kritičan | **Tačno** — :178-179 `hasZbirna/hasOtpremnice` izračunati ali van verdikta; :203 `isValid` bez existence | P1 | U `isValid` uključiti existence kad ima šta da se poredi — zatvara i FM-0013 #7 lanac | S |
| 2 | Recalc vraća True bez post-validacije | Kritičan | **Tačno** — :300-303 commit→True; Complete flow-ovi validiraju posle, Simple/DUPLI grane ne | P2 | Pre commit-a pozvati validator; mismatch → raise (rollback) | S |
| 3 | Sum error postaje validna nula | Kritičan | **Tačno** — EH :93-95, :150-152 vraća nulti/parcijalni dict bez error flag-a; `IsDaFlag` :437-439 nije Null/Error-safe | P1 | `success` ključ u sum dict + `isValid=False` na scan fail | S |
| 4 | Dupli aktivni class redovi | Visok | **Tačno** — :266-279 poslednji aktivni; `ApplyKlasaRecalc` menja samo taj red; validator sabira sve | P2 | Detekcija >1 aktivan red iste klase → raise → MANUAL | S |
| 5 | Stale template iz stornirane generacije | Visok | **Tačno** — :268 prvi red bez obzira na storno → novi class red nasleđuje kupca/vozača/datum stare generacije | P2 | Preferirati aktivan red kao template | S |
| 6 | Semantic headers nisu invariant | Visok | **Dizajnersko ograničenje** — header :7-14 definiše numerički invariant | Prihvaćeno | Dokumentovati obim; opcioni soft-warning | M |
| 7 | Mixed header izvori nisu detektovani | Visok | **Tačno** — `CaptureHeader` :428-435: prva neprazna pobedi | P3 | WARN log na različitu vrstu/sortu izvora | S |
| 8 | Positional insert | Visok | **Tačno** — :349-364 `rowData(0 To 12)` | P2 | Zajednički column-mapped builder sa `SaveZbirna` (uz FM-0011 #2) | M |
| 9 | Monitoring false-negative posle commit-a | Visok | **Netačno** — `Monitor_Event` je ceo pod OERN (modMonitoring.bas:66), ne može podići grešku | Prihvaćeno | Nema akcije | — |
| 10 | Concurrent duplicate class create | Visok | **Delimično** — mehanika tačna, single-writer | P2 | Re-check aktivnog reda klase unutar TX pre append | S |
| 11 | Unknown klasa može proći | Srednji | **Tačno** — :86-88, :143-145: Other ulazi u total | P3 | kgOther/ambOther u poruku + opcioni hard check | S |
| 12 | Decimalna/negativna ambalaža | Srednji | **Tačno** — :69, :128 `CLng` bez validacije | P3 | Data-quality provera u health-check | S |
| 13 | Aktivna prazna zbirna kroz direktan API | Srednji | **Dizajnersko ograničenje** — :230 „0 = zbir ničeg"; storno praznog je odgovornost `RecalcOrStornoEmptyZbirna_TX` | Prihvaćeno | Dokumentovati ugovor | S |
| 14 | Optional header indeks 0 | Srednji | **Tačno** — :334-337 `GetColumnIndex` bez Require | P3 | `RequireColumnIndex` ili guard `c>0` | S |
| 15 | Impact EH daje nepotpun dictionary | Srednji | **Delimično** — :408-411 samo `bothValid`; Dictionary na missing key vraća Empty | P3 | U EH i `oldValid/newValid=False` | S |
| 16 | Case-sensitive business number | Srednji | **Tačno** — :64, :123 binarno poređenje | P3 | `StrComp(..., vbTextCompare)` | S |
| 17 | Mismatch poruka ne opisuje sve uzroke | Srednji | **Tačno** — :455-463 samo KG I/II i AMB | P3 | KG-ukupno i Other grana u poruku | S |
| 18 | Rollback failure nije izolovan | Srednji | **Tačno** — :310 RollbackTx u EH bez OERN | P3 | OERN oko RollbackTx u EH | S |
| 19 | EPS koristi `<`, ne `<=` | Nizak | **Tačno** — :25, :188-190 | P3 | Dokumentovati; praktično zanemarljivo | S |

Bilans: 19 — 14 Tačno, 2 Delimično, 2 Dizajnersko, 1 Netačno (#9). Hitnost: 2×P1 (#1, #3 — hrane FM-0013 lažni COMPLETED), 5×P2, 9×P3, 3×Prihvaćeno.

**Bilans bloka C (91):** 67 Tačno / 9 Delimično / 4 Netačno / 10 Dizajnersko / 1 Nije proverivo. **P1 skup (8):** auto-chain rezultati; `SumByBroj` storno; modStornoFlow #1/#2/#3/#7 (pogrešan class relink, mutacija bez recovery zapisa, lažni COMPLETED ×2); modDokumentInvariant #1/#3 (existence false-valid, fail-open sume). Najisplativije: 6 malih S izmena u modStornoFlow/modDokumentInvariant zatvara sva 4 „lažni uspeh" lanca.

---

## Blok D — Paletni podsistem i frmDokumenta (FM-0016…FM-0018, 67 stavki)

### FM-0016 — `modPaletniList.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | Sledljivost štampa pune otkupe cele zbirne, ne udeo palete | Kritičan | **Tačno** — `GetOtkupiZaPalete` :2196-2308 vraća punu `Kolicina` otkupa preko `TraceByZbirna`; puni :732 i :2691 | P1 | Na listu označiti „mogući izvori" (naslov+fusnota); dugoročno alokacija kg po paletnoj stavci | S–L |
| 2 | Reassign/Detach menjaju prerađenu paletu; prerada ostaje stara | Kritičan | **Tačno** — Reassign :1197-1437 nema `Preradjeno` guard; Detach proverava tek posle skidanja :1704; Adjust blokira :1838 | P1 | `IsPaletaPreradjena` blokada na ulazu Reassign/Detach (kao Adjust) | S |
| 3 | Numeracija može upisati broj 0 (paleta i prerada) | Kritičan | **Tačno** — EH vraća 0 :74 i :2435; bez provere pri insertu :1065 i :2521; okidač redak (RequireSchema prethodi) | P2 | `If broj <= 0 Then Err.Raise` u `CreateNewPaleta` i `SavePrerada_TX` | S |
| 4 | Globalni `mSkipPaletize` može ostati True | Visok | **Tačno** — module-level :31, bez auto-reseta; ali oba callera resetuju i u EH | P2 | Wrapper/parametar umesto globala | M |
| 5 | Dupli class red nove prijemnice tiho pregažen | Visok | **Tačno** — `newById(kl)=...` :1247 poslednji pregazi; isto Adjust :1781 | P2 | Raise na duplikat klase pri mapiranju (obe funkcije) | S |
| 6 | Poslovni broj kao correction root (istorijske generacije) | Visok | **Dizajnersko ograničenje** — selekcija po broju :1288-1297, :1672; broj je jedini root u sistemu | P2 | Koristiti CorrectionID/generation iz `tblStornoVeze` pri selekciji | L |
| 7 | Detach vraća 0 i za no-op i za grešku | Visok | **Tačno** — no-op :1675-1678 = 0; EH :1719-1723 = 0 bez `outInfo` | P2 | EH da vrati -1 (ili raise) + `outInfo` i u EH | S |
| 8 | Zatvorena paleta se automatski reopen-uje | Visok | **Tačno** — :1453-1458 i :2162-2167; reopen nameran („mirror"), ali ručno/zapečaćeno se ne razlikuje | P2 | Kolone `ClosedReason/Sealed`; reopen preskočiti za sealed | M |
| 9 | Nedostajuća tara = 0 bez upozorenja, bruto potcenjen | Visok | **Tačno** — `GetTezinaGajbice` :1174-1176 `NzD`→0; `SpillGajbice` nastavlja :2063 | P2 | Warn/raise kad je `crateW=0` a `tipAmb` neprazan | S |
| 10 | Tehnički ID stavke (`PLS-`) se ne validira | Visok | **Delimično** — :1090-1091 bez provere, ali `GetNextID` nikad ne vraća "" (raise → rollback) | P3 | Defanzivni `If sid = "" Then Err.Raise` | S |
| 11 | Identity lookup nedosledan (strict vs. prvi match) | Visok | **Tačno** — strict :2389-2404 samo u close/prerada; `FindRowIndexByID` u mutacijama :1442, :1463, :1416, :2131 | P2 | `RequireSingleRowIndexByKey` u svim mutacionim lookup-ovima | M |
| 12 | Prerada nema invariant izlaz ≤ ulaz | Visok | **Tačno** — `SavePrerada_TX` :2459-2517 bez provere; komentar :2409 „Bez kalo racunice" | P2 | Blokada/override-warn kad `netoIzlazKg > netoUlaz` | S |
| 13 | Negativne vrednosti pakovanja/težina prolaze | Visok | **Tačno** — samo kombinovani uslov :2469 | P2 | Raise na bilo koju negativnu vrednost na ulazu | S |
| 14 | Partial reassign (orphan klasa) vraća True | Visok | **Delimično** — :1362-1376 warn+commit+True :1431; ugovor dokumentovan, calleri čitaju `outWarn` | P3 | Poseban povratni kod za partial relink | S |
| 15 | Nevalidan `spillMode` tretira se kao PRELIJ | Srednji | **Tačno** — :1931 sve što nije „PREKO" ide u else | P3 | Enum validacija na ulazu (raise na nepoznat mode) | S |
| 16 | Manual close bez reason/seal metapodataka | Srednji | **Tačno** — :527-576 upisuje samo status | P3 | `ClosedReason/At/By` kolone + parametar | M |
| 17 | Fresh undo ostavlja prazne otvorene palete | Srednji | **Tačno** — STEP 1 :1352-1357 ne stornira ispražnjene (Detach to radi :1697-1711) | P3 | U Reassign stornirati ispražnjene ne-prerađene palete kao u Detach | S |
| 18 | Opcione prerada kolone mogu biti tiho preskočene | Srednji | **Tačno** — `EnsurePreradaCols` best-effort :2365-2385; `RequirePreradaSchema` ih ne traži; `PalAppendRow` tiho skipuje :1126 | P2 | Nove kolone dodati u `RequirePreradaSchema` | S |
| 19 | Find-by-number vraća prvi duplikat | Srednji | **Tačno** — :918-935 i :2602-2617; bez dup/storno provere | P3 | Raise na >1 match; preskočiti stornirane | S |
| 20 | Nema location scope-a palete | Srednji | **Dizajnersko ograničenje** — model = 1 workbook po lokaciji | Prihvaćeno | Dokumentovati; kolona Lokacija tek uz deljenje | L |
| 21 | `PrintNepotpunePalete` prekida batch na prvoj grešci | Srednji | **Delimično** — EH prekida petlju :584-610, ALI PRINT/PDF grane gutaju per-item; realan prekid samo PREVIEW | P3 | Per-item `On Error Resume Next` kao `PaletniListOutputClosed` | S |
| 22 | Read schema drift izgleda kao prazni podaci | Srednji | **Dizajnersko ograničenje** — `SafeCell`→Empty :2323-2326; read-modeli namerno fail-soft | P3 | `RequireColumns` u grid read-modelima ili status poruka | S |
| 23 | Sablon vraća `DisplayAlerts` na True, ne na zatečeno | Nizak | **Tačno** — :820-822, EH :913; isto :2759-2761, :2797 | P3 | Sačuvati/vratiti prethodnu vrednost | S |

Bilans: 23 — 17 Tačno, 3 Delimično, 3 Dizajnersko. Hitnost: 2×P1 (sledljivost, prerađena paleta), 11×P2, 9×P3, 1×Prihvaćeno.

### FM-0017 — `frmPalete.frm`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | Strict UI zahteva i kutije i kese; core traži bar jedno | Kritičan | **Tačno** — :262-281 sva 4 uslova blokiraju; core dozvoljava jedno (modPaletniList.bas:2469) | P1 | Uslovna validacija: bar jedna vrsta > 0; tip obavezan samo za korišćenu vrstu | S |
| 2 | Nema prikaza input/output odnosa (ulaz vs. izlaz) | Kritičan | **Tačno** — forma nigde ne sabira neto selektovanih; `RecomputeNeto` :664-677 samo izlaz | P2 | Label „Ulaz Σ / Izlaz / Kalo" + upozorenje kad je izlaz > ulaz | M |
| 3 | Desna lista pogrešno nazvana „Preradjene palete" | Visok | **Tačno** — caption :384; `GetPreradeForGrid` vraća 1 red = cela prerada | P3 | Preimenovati u „Prerade (dvoklik = PDF)" | S |
| 4 | Storno prerade bez preview-a input paleta | Visok | **Tačno** — confirm :566-568 samo broj prerade | P2 | U confirm dodati broj i brojeve paleta iz `tblPreradaStavka` | S |
| 5 | Prazna paleta može biti ručno zatvorena | Visok | **Tačno** — :212-225 bez provere; core ne traži gajbice>0 (modPaletniList.bas:527-576) | P2 | Guard `BrojGajbica > 0` u `ClosePaletaManual_TX` | S |
| 6 | False-success kada `SavePrerada_TX` vrati prazan ID | Visok | **Delimično** — poruka bezuslovna :293-295, ali prazan `preID` praktično nedostižan (raise → rollback) | P3 | Defanzivno: prazan `preID` tretirati kao grešku | S |
| 7 | `VALIDACIJA_UNOSA` toggle isključuje sve UI provere | Visok | **Dizajnersko ograničenje** — :246-281 namerni gejt; OFF → samo core minimum | P2 | Minimalne hard provere (negativne vrednosti, tip za korišćeno pakovanje) van toggle-a | S |
| 8 | Free-text tip kese / gotov proizvod | Visok | **Tačno** — :362-364 `DropDownCombo`; nepoznat tip → `GetTezinaKese`=0 → precenjen neto | P2 | `fmStyleDropDownList` i za kese i za gotov proizvod | S |
| 9 | Nema potvrde pre kreiranja prerade | Visok | **Tačno** — :227-295 direktno snima; undo postoji (storno prerade) | P3 | Yes/No rezime: broj paleta, ulaz Σ, izlaz, proizvod | S |
| 10 | Nema auth/role kontrole u formi | Visok | **Delimično** — ulaz u forme gejtuje `modAuth` (frmOtkupAPP.frm:1073); nema per-akcija prava | P3 | `KorisnikImaPravo` oko storno/prerada akcija | M |
| 11 | Manual close bez razloga | Visok | **Tačno** — :212-225 jedan klik | P3 | Isto kao FM-0016 #16 (`ClosedReason` tok) | M |
| 12 | `ToNum` loše parsira hiljadarske separatore | Srednji | **Tačno** — :680-682 `Val(Replace(",","."))`; „1.234,56"→1.234 | P2 | Reuse postojećeg `TryParseDouble` umesto `ToNum` | S |
| 13 | Batch poruka „poslato na izlaz" i kad je mode OFF | Srednji | **Tačno** — :204-206; core broji i pod OFF | P3 | Pod OFF vratiti 0; poruka razlikuje obrađeno/odštampano | S |
| 14 | `RefreshPrerade` guta greške (stale desna lista) | Srednji | **Tačno** — :524-535 OERN | P3 | LogErr + vidljiva poruka | S |
| 15 | Partial dynamic build ostavlja polufunkcionalnu formu | Srednji | **Tačno** — EH samo LogErr :411-413; retry `Controls.Add` istih imena pod Resume Next :88 | P2 | Cleanup parcijalnih kontrola pre retry + MsgBox | M |
| 16 | Gotov proizvod ostaje izabran posle save-a | Srednji | **Dizajnersko ograničenje** — :694 namerno za serijski rad | Prihvaćeno | Zadržati | S |
| 17 | Nema compatibility provere selektovanih paleta | Srednji | **Tačno** — :229-288 bez provere vrste/sorte; poslovno pravilo nepotvrđeno | P3 | Upozorenje kad selekcija ima >1 vrstu/sortu (bez blokade) | S |
| 18 | Single akcije rade nad focused row u MultiSelect listi | Srednji | **Tačno** — `CurrentPaletaID` :170-173 `ListIndex` | P3 | Upozorenje ako je selektovano više redova | S |
| 19 | Istorija palete se ne prikazuje | Srednji | **Tačno** — grid 13 kolona bez `Istorija` :55-65 | P3 | Prikaz `Istorija` za fokusiranu paletu | M |
| 20 | Nenumerička godina tiho znači „sve godine" | Nizak | **Tačno** — :129-130 `god` ostaje 0 → „sve" | P3 | Poruka „nevalidna godina" | S |

Bilans: 20 — 16 Tačno, 2 Delimično, 2 Dizajnersko. Hitnost: 1×P1 (kutije+kese), 7×P2, 11×P3, 1×Prihvaćeno.

### FM-0018 — `frmDokumenta.frm`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | Novac storno šalje poslovni broj u `NovacID` API | Kritičan | **Tačno** — :3230 `StornoNovac_TX(brDok)`; modStorno.bas:729 očekuje `novacID`; nema lookup-a kao za fakturu :3222 | P1 | Lookup `BrojDokumenta→NovacID` (svi redovi) pre poziva, ili novi `StornoNovacByBrojDok_TX` | S |
| 2 | Stornirane fakture u listi otvorenih | Kritičan | **Tačno** — `FillOpenFakture` :3108-3144 filtrira samo `status <> STATUS_PLACENO`; bez storno provere | P1 | Isključiti stornirane; opciono zahtevati `preostalo > 0` | S |
| 3 | Checkbox određuje da li se Kl.II zbirne validira | Kritičan | **Tačno** — :2216-2221 `kgValid=validKgI` kad je checkbox off; `validKgII` izračunat pa ignorisan; save bez Kl.II :2089 | P1 | Ako izvor ima Kl.II a checkbox off → hard blokada/poruka | S |
| 4 | Smer ambalaže nije obavezan → legacy knjiženje | Kritičan | **Tačno** — :1706 samo `smerCount > 1`; 0 dozvoljeno → `koopSmer=""` → `Case Else` legacy :1303-1308 | P1 | Uz `kolAmb > 0` zahtevati tačno jedan smer | S |
| 5 | Malina auto-zbirna failure je tih | Visok | **Tačno** — :907-912 Resume Next + samo LogErr; korisnik bez upozorenja | P1 | Proveriti Err/rezultat → „zbirna NIJE kreirana — unesite ručno" | S |
| 6 | Immediate correction šalje `allowRelabel=True` bez potvrde | Visok | **Tačno** — :2590 hardkodovano True; recovery panel pita eksplicitno :3975-3987 | P2 | `EvaluatePaletaReassign` + pitanje za RELABEL kao recovery panel | S |
| 7 | Prijemnica correction context nije persistentan | Visok | **Tačno** — `m_pendingRelink*` session-only :33-34; Terminate briše :4508; ostali tipovi imaju `tblStornoVeze` | P2 | Prijemnica-ispravku upisivati u `tblStornoVeze` preko `modStornoContext` | L |
| 8 | Prefill uzima prvu generaciju istog broja | Visok | **Tačno** — :2694-2703 bez `Stornirano`/datum filtera | P2 | Preferirati najnoviju storniranu generaciju | S |
| 9 | Zbirna sa 0 ambalaže blokirana u UI | Visok | **Tačno** — :2248 traži `zbrAmb > 0`; core dozvoljava 0 | P2 | Dozvoliti 0 kad je i suma otpremnica 0 | S |
| 10 | PRIJ recovery bez domain compatibility provere | Visok | **Tačno** — :4005-4025 samo confirm; core proverava samo aktivnost | P2 | U confirm prikazati kupca/vrstu/datum obe strane + warn na mismatch | M |
| 11 | Više aktivnih prijemnica iste zbirne — samo warning | Visok | **Tačno** — :2522-2532 Yes/No override bez razloga | P2 | Uz „Da" tražiti razlog + Monitor event | S |
| 12 | Poslovni broj zavisi od `VALIDACIJA_UNOSA` | Visok | **Tačno** — :1811 (OM) i :2940 (Kupci); prazan broj preskače duplicate check :1647 | P2 | Broj obavezan nezavisno od toggle-a (ili uvek auto-broj) | S |
| 13 | `SaveOMUlaz_TX` živi u UserForm-u | Visok | **Tačno** — :1208-1355 Public TX funkcija u formi | P2 | Premestiti u `modDokumenta` bez izmene logike | M |
| 14 | Detach failure prikazan kao „ništa nije skinuto" | Visok | **Tačno** — :3283-3288 i :4047-4051; koren u core kontraktu (FM-0016 #7) | P2 | Posle core izmene razdvojiti poruke no-op vs. greška | S |
| 15 | Nema auth/role/reason za storno i recovery | Visok | **Delimično** — ulaz u formu gejtuje `modAuth`; nema per-akcija prava/razloga | P3 | `KorisnikImaPravo` za storno/recovery + polje razloga | M |
| 16 | Auto-complete uputstvo vodi u panel bez correction listi | Srednji | **Tačno** — poruka :3582-3585; panel prikazuje samo prijemnice/palete; `GetPendingCorrections` se ne poziva (grep=0) | P3 | Recovery panelu mod „Ispravke" ili preformulisati poruku | M |
| 17 | Correction state se briše i posle neuspelog completion-a | Srednji | **Tačno** — :3612 bezuslovno pre provere `success` | P3 | Brisati session state samo na success | S |
| 18 | „Najnovijih 20 zbirnih" = obrnut fizički red tabele | Srednji | **Tačno** — :1157-1182 reverse iteracija, bez sortiranja po datumu | P3 | Sortirati po Datum (desc) pre uzimanja 20 | S |
| 19 | Live manjak može biti false-green | Srednji | **Tačno** — UI boji po pct :2815-2821; core EH vraća `Array(0,...)` → 0% = zeleno | P2 | Core na grešku vraća marker; UI prikaže „manjak nedostupan" | S |
| 20 | Prosek gajbe uključuje stornirane redove | Srednji | **Tačno** — :2825 → `SumByBroj` bez storno filtera (uz FM-0011 #3) | P2 | Storno-filter u `SumByBroj` | S |
| 21 | Duplicate check-ovi su pre-TX (race window) | Srednji | **Delimično** — :1650, :2066, :2488 pre TX-a; single-writer → nisko | P2 | Re-check duplikata unutar TX-a pre `AppendRow` | S |
| 22 | Runtime paneli mogu ostati parcijalni | Srednji | **Tačno** — `SetupOMIzdavanjeToggle` EH resetuje 2/4 reference :1390-1394 → pojačava #4 | P2 | U EH resetovati sve 4 reference + upozorenje | S |
| 23 | Duplirana provera `kupacID = ""` | Nizak | **Tačno** — :2040-2048 i :2468-2476 (identičan blok 2×) | P3 | Obrisati dupli blok | S |
| 24 | Mešoviti jezici i stari komentari | Nizak | **Tačno** — nem./srp. („ZBIRNA VALIDIERUNG" :2124, „fehlgeschlagen" :1332) | P3 | Ujednačavati postepeno | S |

Bilans: 24 — 22 Tačno, 2 Delimično. Hitnost: 5×P1 (sva 4 Kritična + Malina silent failure), 13×P2, 6×P3.

**Bilans bloka D (67):** 55 Tačno / 7 Delimično / 5 Dizajnersko / 0 Netačno. Najhitniji skup (P1, sve S): frmDokumenta #1–#5, frmPalete #1, modPaletniList #1–#2.

---

## Blok E1 — Novac i banka export (FM-0019…FM-0021, 65 stavki)

### FM-0019 — `modNovac.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | Pozicioni insert 17 kolona može korumpirati ledger pri promeni šeme | Kritičan | **Tačno** — modNovac.bas:197-204, `Array(...)`+`AppendRow` bez `RequireColumns` | P0 | Pre `AppendRow` schema-signature provera (RequireColumns za svih 17) ili upis po imenu kolone | S |
| 2 | Delimični avans menja iznos originalne (bankarske) transakcije | Kritičan | **Tačno** — modNovac.bas:604-605 i 1151 smanjuju original; dizajn-nivo | P2 | Ne dirati bank-sourced red: uvesti append-only alokaciju (`tblNovacAlokacija`) umesto mutacije iznosa | L |
| 3 | Nema allocation lineage (ParentNovacID/SplitFrom) | Kritičan | **Tačno** — split red ima samo napomenu „Avans raspodela" (621, 1167) | P2 | Minimalno: upisati izvorni `NovacID` u napomenu/`OsirocenoOD` split reda | S |
| 4 | Avans kupca može na fakturu drugog kupca | Kritičan | **Tačno** — modNovac.bas:487-490 samo non-empty; 550 lookup bez `KupacID` poređenja | P1 | U `ApplyAvansToFaktura` proveriti `tblFakture.KupacID=kupacID`, inače `Err.Raise` | S |
| 5 | Avans kooperanta može na otkup drugog kooperanta | Kritičan | **Tačno** — modNovac.bas:1096-1098 `FindRows` po ID-u, bez provere kooperanta | P1 | U `ApplyAvansToOtkup` proveriti `COL_OTK_KOOPERANT=kooperantID`, inače `Err.Raise` | S |
| 6 | Avans može na stornirani otkup/fakturu | Kritičan | **Tačno** — modNovac.bas:1094-1098 sirov `GetTableData` bez storno provere; faktura isto (550) | P1 | Guard: target `Stornirano="DA"` → `Err.Raise` u oba apply-a | S |
| 7 | Nema document-level novac storna (broj ≠ NovacID) | Kritičan | **Tačno** — modul nema resolver; frmDokumenta.frm:3230 šalje `brDok` kao `NovacID` | P1 | Resolver `BrojDokumenta→aktivni NovacID(i)` + izbor stavke u frmDokumenta | M |
| 8 | `DatumIsplate` = `Date` obrade, ne datum isplate | Visok | **Tačno** — modNovac.bas:883 (uz očuvanje postojećeg, 882) | P2 | Upisivati max `Datum` aktivnih isplata umesto `Date` | S |
| 9 | Reset odvaja SVE isplate, direktna postaje orphan | Visok | **Tačno** — modNovac.bas:1248-1261 bez tip-filtera; storno tok modStorno.bas:177/1141 isto | P2 | Ograničiti unlink na `NOV_VIRMAN_AVANS_KOOP`; za ostale postaviti `OsirocenoOD` | S |
| 10 | `OsirocenoOD` postoji ali se ne postavlja | Visok | **Tačno** — SaveNovac upisuje "" (201); nijedan setter za tblNovac u repou | P2 | `ResetNovacOtkupLink` da upiše `otkupID` u `OsirocenoOD` | S |
| 11 | Apply wrapper vraća `True` za no-op | Visok | **Tačno** — modNovac.bas:1198-1204; nepostojeći target `Exit Sub` (1098) | P1 | `_TX` da vraća primenjeni iznos (ByRef); nepostojeći target → greška | S |
| 12 | Core dozvoljava prazan broj dokumenta | Visok | **Tačno** — `ValidateNovacInput` (1347-1386) ne proverava `brojDok` | P2 | Obavezan `brojDok` za tipove koji ga poslovno zahtevaju | S |
| 13 | Tip-specifična pravila (smer, obavezna polja) nisu validirana | Visok | **Tačno** — validacija samo neprazan tip (1357-1360) | P2 | `Select Case tip` → obavezni `OMID`/`FakturaID`/`OtkupID` + smer | M |
| 14 | Raspodela po fizičkom redu tabele, ne FIFO po datumu | Visok | **Tačno** — petlje 560/1115; komentar „chronologisch" (558) netačan | P2 | Sortirati kandidate po `Datum`+`NovacID` pre raspodele | S |
| 15 | Multi-product faktura pripisana samo prvoj vrsti | Visok | **Tačno** — `BuildVrstaFakturaCache` prva stavka pobedi (783-789) | P2 | Raspodela po vrednosti stavki ili kategorija „mešovita faktura" | M |
| 16 | `GetOpenOtkupi` veruje stale statusu `Isplaceno` | Visok | **Tačno** — 974/1001 status pre računa; realan trigger: `StornoNovac` (modStorno.bas:781-785) ne zove `UpdateOtkupStatus` | P1 | U `StornoNovac` čitati `OtkupID` i pozvati `UpdateOtkupStatus` (+ snapshot tblOtkup) | S |
| 17 | Nema cross-user allocation lock-a | Visok | **Delimično** — tačno, ali single-writer desktop; `clsTransaction` lokalni snapshot | P2 | Opciono: version/ModifiedAt provera pre mutacije | M |
| 18 | `GetBankaByPartner` filtrira po nazivu, ne ID-u | Srednji | **Tačno** — modNovac.bas:33-40 `COL_NOV_PARTNER` | P2 | Filtrirati po `PartnerID`, naziv kao fallback | S |
| 19 | Partner mapa bez TX/istorije, positional append | Srednji | **Tačno** — 276-331, `Array` 4 kolone bez TX | P2 | Append po imenu kolona + datum/korisnik kolone | S |
| 20 | `SaveNovac_TX` snapshotuje više nego što menja | Srednji | **Tačno** — 71-73 tri tabele, piše samo tblNovac | P3 | Ukloniti nepotrebne snapshote ili dokumentovati ugovor | S |
| 21 | `GetUplataByVrsta` ne filtrira tip novca | Srednji | **Delimično** — tačno (412-417), ali ID-prefiksi razdvajaju kupce/kooperante — kolizija malo verovatna | P3 | Dodati filter `EntitetTip="Kupac"` | S |
| 22 | `GetUplataForOtkup` pogrešan naziv (sabira Isplatu) | Srednji | **Tačno** — 796-797 komentar priznaje; alias postoji (894-896) | P3 | Pozivaoce preusmeriti na `GetIsplataForOtkup`, stari označiti deprecated | S |
| 23 | Statusi su mutable cache bez invariant heal-a | Srednji | **Tačno** — heal samo kroz eksplicitne `UpdateOtkupStatus` pozive; storno gap (v. #16) | P2 | Pozvati recompute na svim mutacionim tokovima (storno pre svega) | S |

Bilans: 23 — Tačno 21, Delimično 2, Netačno 0; hitnost: P0 1, P1 6, P2 13, P3 3.

### FM-0020 — `frmBankaExportPregled.frm`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | Stale override posle reload-a — CSV može preći novi otvoreni saldo | Kritičan | **Tačno** — `PruneStaleOverrides` (504-535) briše samo nestale ID-eve, bez clamp-a; `CollectIsplataBlokovi` (967) ne revalidira | P1 | U `PruneStaleOverrides` clamp override na novi `OtvorenIznos`; revalidacija i u collect-u | S |
| 2 | Batch avans `True` za no-op — naduvan `okCount` | Kritičan | **Tačno** — frm:1104 broji svaki `True`; core no-op vraća `True` (modNovac 1098/1204) | P1 | `ApplyAvansToOtkup_TX` da vraća primenjeni iznos; brojati samo iznos>0 | S |
| 3 | Nema selekcije = svi prikazani blokovi | Visok | **Tačno** — frm:947-956 + komentar 943; status istovremeno „Nista nije selektovano" (859) | P2 | Bez selekcije → hard stop + dugme „Izaberi sve filtrirane" | S |
| 4 | Export rezervacija nije potvrđena — moguć ponovni izvoz | Visok | **Tačno** — nikakav export status/trag ne postoji (dizajn „knjiži se iz izvoda") | P2 | Append-only export log + upozorenje pri ponovnom izvozu istog `OtkupID` | M |
| 5 | Batch avans bez preview-a raspodele | Visok | **Tačno** — frm:1095-1097 samo broj blokova | P2 | Dry-run simulacija raspodele po redosledu prikaza pre potvrde | M |
| 6 | Single avans može false-success (stale + no-op True) | Visok | **Tačno** — frm:1061-1063 „Avans vezan" na svaki `True` | P2 | Isto kao #2: prikaz stvarno primenjenog iznosa | S |
| 7 | Payer račun se resetuje na Activate | Visok | **Tačno** — frm:48-50 → `PopulateRacunCombo` vraća default (339-348) | P2 | Pre `Clear` zapamtiti izbor i restaurirati ako račun i dalje postoji | S |
| 8 | Nema auth/role kontrole | Visok | **Dizajnersko ograničenje** — aplikacija nema role model uopšte | Prihvaćeno | — | — |
| 9 | UI opis avansa prikriva split mutaciju | Visok | **Tačno** — frm:1057 „OtkupID na avans red" vs split u modNovac (1151-1168) | P3 | Dopuniti tekst potvrde: delimičan avans deli red na dva | S |
| 10 | Override se briše i kod no-op `True` | Visok | **Tačno** — frm:1023-1026 | P2 | Uklanjati override samo ako se `OtvorenIznos` bloka stvarno promenio | S |
| 11 | PDF nema završnu potvrdu | Srednji | **Tačno** — `btnExport_Click` (898-930) bez MsgBox potvrde | P3 | Ista Yes/No potvrda (count/sum) kao CSV | S |
| 12 | PDF dozvoljava prazan payer račun | Srednji | **Delimično** — frm:921 ne proverava, ali modul ima `SELLER_ACCOUNT` fallback (mod 234-235) | P3 | Ista provera `SelectedRacun` kao u CSV grani | S |
| 13 | Datum filter fail-open (`On Error Resume Next`) | Srednji | **Tačno** — frm:454-457 | P2 | TryParse + poruka + prekid učitavanja na nevalidan datum | S |
| 14 | Runtime build se ne retry-uje (`m_SetupDone` pre setup-a) | Srednji | **Tačno** — frm:53 pre `EnsureRuntimeControls` (57); helperi log-only | P3 | Zvati `EnsureRuntimeControls` i na reaktivaciji (idempotentan je) | S |
| 15 | Nema sačuvanog draft paketa | Srednji | **Dizajnersko ograničenje** — override je svesno in-memory draft | Prihvaćeno | — | — |
| 16 | Nema user/batch audita u formi | Srednji | **Tačno** — nijedan `Monitor_Event` u formi (ni u modulu) | P2 | `Monitor_Event` pri CSV/PDF: count, sum, račun, putanja | S |
| 17 | Tekući račun samo Boolean guard | Srednji | **Tačno** — forma veruje `HasTekuciRacun`; builder = samo neprazan (mod 132) | P2 | Format/mod-97 kontrola računa u builderu + upozorenje | S |
| 18 | Dupli `OtkupID` ruši load | Srednji | **Tačno** — frm:518 `Add` bez `Exists`; trigger zahteva dupli ID u tblOtkup | P3 | `If Not currentSet.Exists(...) Then Add` | S |
| 19 | Datum valute fiksno danas | Srednji | **Tačno** — frm:1220; generator isto (mod 357-358) | P3 | Opciono polje „Datum valute" → parametar `GenerisiNalogeCSV` | S |
| 20 | Prazan kooperant combo može error (`LBound/UBound`) | Nizak | **Netačno** — `Keys` praznog Dictionary = niz (0 To -1); `LBound/UBound` legalni, obe petlje se preskaču bez greške (frm:383-405) | — | Ništa | — |
| 21 | `Caption = UserForm1` u source-u | Nizak | **Tačno** — frm:3; chrome se ionako skida | P3 | Postaviti Caption u Activate (kozmetika) | S |

Bilans: 21 — Tačno 17, Delimično 1, Netačno 1 (#20), Dizajnersko 2; hitnost: P1 2, P2 9, P3 7, Prihvaćeno 2.

### FM-0021 — `modBankaExportPregled.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | Nema finalne saldo revalidacije — CSV može sadržati preplatu | Kritičan | **Tačno** — mod:369 proverava samo `HasTekuciRacun` i `IsplatitiIznos>0`; potvrđuje FM-0020 #1 | P1 | Pre upisa sveže `BuildBlokIsplataList`/saldo po `OtkupID`; iznos > otvoreno → prekid | M |
| 2 | Multi-class redovi dele isti poziv na broj | Kritičan | **Delimično** — tačno (mod:378 upisuje `BrDok`; klase = zasebni redovi), ali import odbija ambigvitet (`hitCount=1`, modBankaMapiranje.bas:1423-1425) i pada na ručno — ne knjiži pogrešno | P2 | Agregirati klase istog `BrDok` u jedan nalog (zbir) ili sufiks klase u poziv | M |
| 3 | Nema export ledger/rezervacije — isti blok više puta izvoziv | Kritičan | **Tačno** — samo fajl (390); header komentar 20-21; bez ikakvog traga | P2 | `tblBankaNalogLog` (append-only) + upozorenje na ponovni export istog bloka | M |
| 4 | Poslovni broj kao jedini bankarski correlation ključ | Kritičan | **Delimično** — resolver je skopiran po kooperantu + unique-hit + račun cross-check (modBankaMapiranje.bas:1412-1425); kolizija završi kao ručno mapiranje | P3 | Dokumentovati ograničenje; po potrebi jedinstveni token u pozivu | M |
| 5 | Otvoren red bez `KooperantID` tiho nestaje | Visok | **Tačno** — mod:100-102 `GoTo NextRow` bez traga | P2 | Brojati preskočene i prikazati u `SummarizeBlokList`/statusu | S |
| 6 | Nevalidan datum tiho uklanja obavezu | Visok | **Tačno** — mod:81-92 skip bez poruke | P2 | Isti brojač preskočenih + prikaz razloga | S |
| 7 | Tekući račun samo neprazan | Visok | **Tačno** — mod:132; `NormalizujRacun` samo skida razmake (419-421) | P2 | NBS format/mod-97 provera; nevalidan račun → blokiran blok | S |
| 8 | Nema export audita (ko/kada/koliko) | Visok | **Tačno** — samo `LogErr` na exception (287, 396) | P2 | `Monitor_Event` za CSV/PDF sa count/sum/račun/path | S |
| 9 | Nema auth/approval kontrole | Visok | **Dizajnersko ograničenje** — app nema role model | Prihvaćeno | — | — |
| 10 | Stale `Isplaceno` može sakriti dug | Visok | **Tačno** — dokazan uzrok: `StornoNovac` ne zove `UpdateOtkupStatus` (modStorno.bas:781-785) | P1 | Fix u `StornoNovac` (čitaj `OtkupID` → `UpdateOtkupStatus`, + snapshot tblOtkup) | S |
| 11 | Ime fajla sekundna rezolucija — moguće prepisivanje | Visok | **Tačno** — `hhnnss` (281, 389); writer prepisuje; ručni tok praktično sprečava | P3 | Sekvenca + collision check pre upisa | S |
| 12 | Datum valute fiksno danas | Visok | **Tačno** — mod:357-358; svesno ograničenje | P3 | Parametar datuma valute (v. FM-0020 #19) | S |
| 13 | PDF nema success rezultat (`Sub`) | Srednji | **Tačno** — mode OFF bez izlaza (283), `ws Is Nothing` → tihi izlaz (271) | P3 | Function koja vraća putanju/False; forma proverava | S |
| 14 | Računi firme nisu validirani/deduplikovani | Srednji | **Tačno** — mod:302-317 samo spajanje | P3 | Trim+dedup+format provera u `BankaNalogRacuniCSV` | S |
| 15 | Jedan generički CSV za sve banke | Srednji | **Nije proverivo statički** — kod potvrđuje jedan format (361-380); prihvatljivost po banci van koda | P3 | Operativno potvrditi sa svakom bankom | M |
| 16 | Nema eksplicitnog sortiranja liste | Srednji | **Tačno** — mod:139 redom `GetOpenOtkupi` | P3 | Sort po `Datum`+`BrDok` | S |
| 17 | Dupli `KooperantID` — prvi TR pobedi bez prijave | Srednji | **Tačno** — mod:171-173; trigger = korumpirani matični podaci | P3 | Logovati duplikat pri build-u cache-a | S |
| 18 | Model poziva na broj prazan | Srednji | **Tačno** — mod:377; prihvatljivost banke nije statički proveriva | P3 | Config `BANKA_NALOG_MODEL` (prazno = današnje) | S |
| 19 | CSV `""` rezultat višeznačan | Srednji | **Tačno** — 342-343, 351, 385, 397 | P3 | ByRef outReason ili struktura rezultata | S |
| 20 | File-write atomicitet nije potvrđen | Srednji | **Tačno** — `WriteAllTextUtf8` = ADODB `SaveToFile` overwrite, bez temp+rename | P3 | Temp fajl pa `Name` rename | S |
| 21 | Dokumentacija kaže 3 računa, kod koristi 4 | Nizak | **Tačno** — komentari mod:294-295 i frm:239 vs kod `RACUN_1..4` (305-306) | P3 | Ispraviti komentare na „1..4" | S |

Bilans: 21 — Tačno 17, Delimično 2, Dizajnersko 1, Nije proverivo 1; hitnost: P1 2, P2 6, P3 12, Prihvaćeno 1.

**Bilans bloka E1 (65):** Tačno 55 / Delimično 5 / Netačno 1 / Dizajnersko 3 / Nije proverivo 1. Ključno: P0 pozicioni `SaveNovac`; P1 lanac preplate (stale override + bez finalne revalidacije); P1 `StornoNovac` ne osvežava `Isplaceno` status (sakriven dug); P1 target-owner/active/no-op guardovi avansa.

---

## Blok E2 — Banka import i mapiranje (FM-0022…FM-0024, 74 stavke)

Prefiksi u dokazima: `mBM:` = modBankaMapiranje, `frm:` = frmBankaImport.

### FM-0022 — `modBankaImport.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | Dedupe ključ ne uključuje broj računa firme | Kritičan | **Tačno** — :783-798: BrojDok+Referenz ili datum/uplata/isplata/partner; BrojRacuna nigde; izvod-brojevi se sudaraju preko računa → tihi drop transakcije | P1 | U `IsDuplicateBankaImport` dodati poređenje `COL_BIM_BROJ_RACUNA` u obe grane | S |
| 2 | Nema cross-user unique/lock zaštite | Kritičan | **Delimično** — :49-104 TX lokalni, ali Excel file-lock daje single-writer; ista instanca zaštićena dedupe-om | P2 | Dokumentovati single-writer; opciono import-lock flag u LocalConfig | S |
| 3 | Nepoznata banka pada na Komerc parser | Visok | **Tačno** — :363-364 `Else → "KOMERC"`; ublaženo: 4-nivo integrity obara pogrešan parse | P2 | `DetectBank` vrati `UNKNOWN` → hard greška pre dispatch-a | S |
| 4 | Poison PDF rollbackuje ceo batch | Visok | **Dizajnersko ograničenje** — :162-168 + EH :80-103; nameran all-or-nothing (komentar :12-14) | P2 | Orchestrator „jedan PDF = jedan TX" oko postojećeg `_Core` | M |
| 5 | Post-commit move failure izgleda kao import failure | Visok | **Tačno** — :74-77: posle `CommitTx` move greška ide u isti EH i `Err.Raise` :103, a DB je trajan | P2 | Posle commita lokalni handler: WARN „uvoz OK, fajl nije premešten" | S |
| 6 | Copy+delete može duplirati fajl | Visok | **Tačno** — :991-1001 Copy→Delete; retry pravi `_001` kopiju; podaci zaštićeni dedupe-om | P3 | Pri retry preskočiti copy ako target iste veličine postoji | S |
| 7 | Stornirani raw red omogućava reimport | Visok | **Delimično** — :768 `ExcludeStornirano` pre dedupe; verovatno nameran recovery, fali audit veza | P3 | Dokumentovati kao recovery | S |
| 8 | Fallback dedupe exact-match je slab | Visok | **Tačno** — :791-794 exact `Double` + string datum + trim partner; identične no-ref transakcije istog izvoda → druga tiho ispuštena | P2 | U fallback ključ dodati redni broj transakcije u izvodu | M |
| 9 | Nema ImportBatch/IzvodID | Visok | **Tačno** — :549-569, :632-654: header ID ne postoji | P2 | Izvedeni `BankaIzvodID` (BrojRacuna+BrojIzvoda+godina) kao staging kolona | M |
| 10 | Valuta hardkodovana RSD | Visok | **Tačno** — :688; rizik uslovan uvođenjem deviznog računa | P3 | Parsirati valutu ili hard-fail na ne-RSD | S |
| 11 | Mapping nije deo import TX-a | Visok | **Dizajnersko ograničenje** — dvostepena arhitektura namerna | Prihvaćeno | Korelacija preko budućeg IzvodID (vidi #9) | — |
| 12 | Error-move failure blokira svaki sledeći batch | Visok | **Tačno** — EH :85-93 Resume Next; fajl ostaje u Inbox → ponovni pad | P2 | Skip-lista problematičnih fajlova / quarantine folder | M |
| 13 | Dedupe je O(n×m) | Srednji | **Tačno** — :761 `GetTableData` po svakom redu | P3 | Dictionary dedupe ključeva jednom po batch-u | M |
| 14 | GetNextID po redu | Srednji | **Tačno** — :669 u petlji; single-writer → samo performanse | P3 | Ništa hitno | S |
| 15 | Error klasifikacija može pogrešiti | Srednji | **Tačno** — :856-858 APPENDROW grana pre SCHEMA; re-raise :737 menja source | P3 | SCHEMA granu iznad APPENDROW | S |
| 16 | File moves nisu per-item izolovani | Srednji | **Tačno** — :898 proceduralni EH prekida petlju na prvoj grešci | P3 | Per-item Resume Next + zbirna greška | S |
| 17 | Datum se čuva kao string | Srednji | **Tačno** — :680, :682 `CStr`; downstream `CDate` locale-zavisan (mBM:263) | P2 | Typed `Date` u staging, string za audit | M |
| 18 | Drive file order nedefinisan | Srednji | **Tačno** — :1360-1372 `Dir$` + maxFiles :1255; bitno tek kod >50 fajlova | P3 | Sortirati kolekciju po imenu | S |
| 19 | Readiness double-size check prekratak | Srednji | **Tačno** — :1386-1395 samo `DoEvents`; ublaženo copy-size proverom :1316-1321 | P3 | Kratak Sleep (300–500 ms) između merenja | S |
| 20 | Nema success batch monitoringa | Srednji | **Tačno** — grep `Monitor_` = 0; samo `Debug.Print` :718 | P3 | `Monitor_Event "BANKA_IMPORT_SUMMARY"` | S |
| 21 | Nema auth/period kontrole | Srednji | **Dizajnersko ograničenje** — desktop sa jednim operaterom | P3 | Ništa sada | — |
| 22 | Dupliran filename helper | Nizak | **Tačno** — :823-831 `GetFileNameFromPath2` = kopija :952-961 u ISTOM modulu; + `NzBIM` dupliran vs mBM:1799 | P3 | Obrisati `GetFileNameFromPath2` (poziv :1163 → original) | S |
| 23 | Veliki dijagnostički kod u produkciji | Nizak | **Prihvaćeno** — :369-407, :1053-1113, :1124-1181; Alt+F8 dijagnostika je praksa projekta | P3 | Opciono izdvojiti u `modBankaDiag` | S |

Bilans: 23 — Tačno 17, Delimično 2, Dizajnersko 3, Prihvaćeno 1. Hitnost: 1×P1 (multi-account dedupe — jedini tihi gubitak podataka), 8×P2, ostalo P3.

### FM-0023 — `modBankaMapiranje.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | Isti BIM može biti knjižen dvaput (nema CAS/lock) | Kritičan | **Delimično** — mapped-status guard POSTOJI: `ValidateBankaImportNotProcessed` :1590-1637 čita sveže stanje i odbija Da/Skip/storno; single-thread + file-lock → race praktično nedostižan | P2 | Dokumentovati; opciono re-check pre `UpdateBankaImportStatus` | S |
| 2 | Kupac/faktura owner nije validiran | Kritičan | **Tačno** — :256-259 samo `RequireSingleRow` (postojanje), bez `Faktura.KupacID=kupacID` i bez storno provere; UI šalje uvek `""` (frm:332), auto put owner-konzistentan → latentno | P2 | U `MapBankaImportAsKupac` provera `COL_FAK_KUPAC` + nije stornirana | S |
| 3 | Kooperant/otkup owner nije validiran | Kritičan | **Tačno** — :400-403 i `LinkNovacToOtkupStrict` :2126-2149 samo postojanje; block-putevi izvode target iz kooperantID → ekspozicija samo test/API | P2 | U `MapBankaImportAsKooperant` provera `COL_OTK_KOOPERANT` + storno | S |
| 4 | Multi-class lineage nije očuvan | Kritičan | **Delimično** — greedy „veći saldo prvi" :1500-1513, :837-841; poziv nosi samo BrDok (bez klase) → informacija ne postoji u banci; total konvergira | P2 | Exact-amount match pre greedy | S |
| 5 | `BrDok` nema generation/station scope | Visok | **Tačno** — :1459-1462 samo kooperant+BrDok; zahteva reuse broja + stari otvoren red | P2 | Upozorenje/filter za kandidate sa dalekim datumima/sezonama | M |
| 6 | Non-TX public API može ostaviti parcijalno stanje | Visok | **Tačno** — sve Map* bez `_TX` su Public (:221, :356, :500…); spolja ih niko ne zove | P2 | Non-TX mutatore na `Private` | S |
| 7 | KupacID/OMID ne moraju postojati | Visok | **Tačno** — :234-248 (naziv-fallback :248), :512-526 (`omNaziv=omID`); UI put guarded → API rupa | P2 | `RequireSingleRow` nad TBL_KUPCI/TBL_STANICE | S |
| 8 | Ručni mapping ne proverava smer | Visok | **Tačno** — nijedna Map* ne traži smer; sirovi iznosi :272-273, :416-417, :539-540; **dostižno iz UI** (cmbMapTip slobodan) | P1 | Guard smera po tipu (Kupac: uplata>0 ∧ isplata=0; obrnuto kooperant) | S |
| 9 | OM oba smera pod istim `NOV_KES_FIRMA_OTKUPAC` | Visok | **Tačno** — :538-540 | P2 | Poseban tip za bankarski OM promet ili smer-guard | M |
| 10 | Invoice fallback bez preostalog salda/statusa | Visok | **Tačno** — :1332/:1353 `colStatus` učitan pa neiskorišćen; :1379 match na PUNU vrednost fakture | P2 | `TryResolveFakturaForKupac`: preostali saldo + preskočiti plaćene | M |
| 11 | Nema strukturne BIM→Novac veze | Visok | **Tačno** — samo `BuildBIMNapomena` tekst :1766-1782 | P2 | Kolona `BankaImportID` u tblNovac + upis pri SaveNovac | M |
| 12 | Storno Novac ne otvara BIM za remap | Visok | **Tačno** — reverse workflow ne postoji | P2 | U storno putu modNovac: `BIM:` marker → reset `Obradjeno=""` | M |
| 13 | Nema kandidata → cela isplata avans | Visok | **Tačno** — :793-818: avans + status `Da`, bez Manual-Required; maskira pogrešan poziv | P2 | Kad poziv postoji a kandidata 0 → ostaviti otvoreno; avans samo ručno | M |
| 14 | Više od 2 kandidata → runtime failure | Visok | **Tačno** — :1457 `ReDim 1 To 2`; treći pogodak → Subscript out of range :1492-1497 → rollback (u AutoMapAll obara ceo batch) | P1 | `count>2` → `Err.Raise` sa jasnom porukom pre kopiranja | S |
| 15 | Skip nema razlog/by/at | Visok | **Tačno** — :904-940; userId `"Operator"` :933 | P2 | Obavezan razlog → monitoring poruka | S |
| 16 | Nema auth/locked-period kontrole | Visok | **Dizajnersko ograničenje** — desktop, jedan operater | P3 | Ništa sada | — |
| 17 | PartnerMap failure se ignoriše | Srednji | **Tačno** — `Call savePartnerMap` :285, :429, :552, :812, :896 — rezultat odbačen | P3 | `If Not savePartnerMap(...) Then LogErr` | S |
| 18 | Kupac account resolver ne filtrira lifecycle | Srednji | **Delimično** — :1836-1846 bez `ExcludeStornirano` (kooperant ima :1872); storno kolona u tblKupci nepotvrđena | P3 | Dodati `ExcludeStornirano` (no-op kad kolone nema) | S |
| 19 | NormalizeKonto nije checksum validacija | Srednji | **Tačno** — :1971-2004; nema dužine 18 ni mod-97 | P3 | WARN ako rezultat nema 18 cifara | S |
| 20 | `Error` status bez koda/poruke | Srednji | **Tačno** — :1639-1655 samo literal | P3 | Razlog poslednjeg pokušaja u monitoring/napomenu | M |
| 21 | Batch hard failure rollbackuje sve | Srednji | **Tačno** — jedan TX :1007-1025; „MappedBeforeFail" :1071 broji rollbackovane (zbunjujuće) | P2 | Per-red TX u AutoMapAll + tačan summary | M |
| 22 | Monitoring hardkoduje `Operator` | Srednji | **Tačno** — :933, :996, :1034, :1073, :2082, :2118 | P3 | `Environ("USERNAME")` helper | S |
| 23 | Datum locale-sensitive string→CDate | Srednji | **Tačno** — `CDate(bim(1,2))` :263, :407, :530, :845, :875; ista mašina → praktično konzistentno | P2 | Rešava se typed datumom u stagingu (FM-0022 #17) | M |
| 24 | MsgBox u business modulu | Srednji | **Tačno** — :133, :160, :235… + `gBankaSilentBatch` :59 | P3 | Vremenom rezultat-objekat; silent flag zadržati | M |
| 25 | Max-2 kao implicitni model | Nizak | **Tačno** — :1457 hardkod + komentar :1500 | P3 | `Const MAX_BLOCK_CLASSES = 2` + jasna greška (uz #14) | S |

Bilans: 25 — Tačno 21, Delimično 3, Dizajnersko 1. Hitnost: 2×P1 (#8 smer dostižan iz UI; #14 stvaran runtime pad), 12×P2, ostalo P3.

### FM-0024 — `frmBankaImport.frm`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | Otvaranje forme knjiži novac (Activate side-effect) | Kritičan | **Tačno** — :70-74: `AutoMapStrongKeysBankaImport_TX` u Activate, bez potvrde/preview-a, bez feedback-a | P1 | Na Activate samo izračunati kandidate; „Strong-map: N [Primeni]" → knjiži na klik | M |
| 2 | Manual Kupac uvek postaje avans | Kritičan | **Tačno** — :332 `MapBankaImportAsKupac_TX(bimID, kupacID, "", True)` uvek `""` → `NOV_KUPCI_AVANS`, dok preview može prikazati fakturu (:502-505) | P1 | Završiti `cmbFaktura` i proslediti izbor; do tada upozorenje „ručno = avans" | M |
| 3 | Manualni blok: preview koristi drugi izvor | Kritičan | **Tačno** — :249-251 combo → preview, ali `BuildOutgoingPreview` čita poziv iz tabele (:595, :622); komanda koristi `cmbOtkupBlok` (:340-346) | P1 | U preview za tip Kooperant koristiti `cmbOtkupBlok.value` ako je popunjen | S |
| 4 | `Auto sve` bez potvrde | Visok | **Tačno** — :311-316 odmah pokreće heuristički batch | P2 | `vbYesNo` sa brojem otvorenih + sumama | S |
| 5 | Strong auto-map failure potpuno tih | Visok | **Tačno** — :72-74 Resume Next + ignorisan return | P2 | Rezultat u `lblStatus`; greška → upozorenje | S |
| 6 | Ručni mapping bez finalne potvrde | Visok | **Tačno** — :318-355 direktno knjiži | P2 | Confirm rezime (BIM, tip, target, iznos) | S |
| 7 | Kupac/OM koriste naziv kao identitet | Visok | **Tačno** — :331, :350 `LookupValue` po Nazivu (prvi pogodak); kooperant ima ID-display | P2 | „ID - Naziv" + `ExtractIDFromDisplay` i za kupce/OM | S |
| 8 | Target liste bez lifecycle filtera | Visok | **Tačno** — :227 kooperanti bez `ExcludeStornirano`; kupci/OM preko lookup-a bez filtera | P2 | `ExcludeStornirano` u `LoadManualTargets` | S |
| 9 | Otkup combo uključuje zatvorene blokove | Visok | **Tačno** — :252-290 bez saldo provere → zatvoren blok = ceo iznos avans bez upozorenja | P2 | Samo blokovi sa otvorenim saldom ili oznaka „(zatvoren)" | M |
| 10 | Skip jedan klik bez potvrde/razloga | Visok | **Tačno** — :357-370 | P2 | Potvrda + razlog (uz FM-0023 #15) | S |
| 11 | Return rezultati manual mappinga ignorisani | Visok | **Tačno** — :332 `Call`, :343-345 `n` neiskorišćen, :351 `Call` | P2 | Proveriti rezultat, prikazati uspeh/failure | S |
| 12 | Multi-account identitet nije prikazan | Visok | **Tačno** — grid :111 i detalji :187-198 bez računa/banke/izvoda | P2 | Dodati BrojRacuna (+ banku) u detalj panel | S |
| 13 | Jak ključ partner-konto nije vidljiv | Visok | **Tačno** — konto u resolveru (:461, :594) a ne prikazuje se | P2 | `lblKonto` u detaljima | S |
| 14 | Nema pregleda nastalih NovacID | Visok | **Tačno** — posle akcije samo `LoadBankaRows` | P2 | Poruka sa NovacID/brojem kreiranih redova | S |
| 15 | Nema auth/period kontrole | Visok | **Dizajnersko ograničenje** — kao FM-0022 #21 / FM-0023 #16 | P3 | Ništa sada | — |
| 16 | Preview ne prikazuje split raspodelu | Srednji | **Tačno** — :638-645 kandidati + otvoreno, bez alokacije | P3 | Simulirati greedy raspodelu u preview tekstu | M |
| 17 | Preview može zastareti | Srednji | **Delimično** — bez verzije/timestampa, ali single-writer + core revalidacija → teorijski | P3 | Ništa hitno | — |
| 18 | Latest statement arbitraran kod multi-account | Srednji | **Tačno** — :793-815 max datum, tie → poslednji fizički red | P3 | BrojRacuna u summary; kasnije grupisanje po računu | S |
| 19 | KPI Uplate/Isplate dvosmislene | Srednji | **Tačno** — :906-936 sumira samo otvorene, naslovi generički | P3 | Naslovi „Otvorene uplate/isplate" | S |
| 20 | Mapirano denominator meša Skip/Error | Srednji | **Tačno** — :944-966 | P3 | Razdvojiti prikaz (Da / Open / Error / Skip) | S |
| 21 | Datum format može oboriti ceo load | Srednji | **Delimično** — :159 bez per-row EH, ali `Format$` ne diže grešku; realniji pad je `CDbl` :162-163 | P3 | Lokalni `On Error` u `LoadBankaRows` | S |
| 22 | AddItem punjenje slabije skalira | Srednji | **Tačno** — :157-167; backlog tipično mali | P3 | Array assignment kad zatreba | M |
| 23 | Header build fail-soft | Srednji | **Dizajnersko ograničenje** — :102 Resume Next, namerno idempotentno | P3 | Ništa | — |
| 24 | Auto-jedan failure uglavnom tih | Srednji | **Tačno** — :304-306 poruka samo na uspeh | P3 | `Else` grana: „Nije mapirano — proveri preview/status" | S |
| 25 | Caption `UserForm1` u source-u | Nizak | **Tačno** — :3; runtime naslov preko `lblKopf` | P3 | `Me.Caption` u Activate (runtime) | S |
| 26 | `cmbFaktura` mrtva kontrola | Nizak | **Tačno** — :216 Clear, nikad punjena; koren nalaza #2 | P3 | Implementirati u sklopu #2 | M |

Bilans: 26 — Tačno 22, Delimično 2, Dizajnersko 2. Hitnost: 3×P1 (sva tri Kritična potvrđena), 11×P2, ostalo P3.

**Bilans bloka E2 (74):** 60 Tačno / 7 Delimično / 7 Dizajnersko-Prihvaćeno. Ključno: dedupe bez broja računa (P1, tihi gubitak transakcije u multi-account radu); subscript pad kod 3+ kandidata; smer ručnog mapiranja dostižan iz UI; tri preview/command neslaganja u formi. „Cross-user" kritični su Delimično/P2 — mapped-status guard postoji.

---

## Blok F — Izveštaji i pomoćne klase (FM-0025…FM-0029, 110 stavki)

**Ključna otkrića koja koriguju FM težine:** (1) jedini caller `OutputToSheet` čisti sheet pre upisa (`modPrint.bas:42`) — stale-report rizik je latentan; (2) `GroupBySum`/`SumColumn` i enum `IzvestajTip` su mrtav kod (0 pozivalaca); (3) `ReportManjak` OM-kombinacija nije dostižna iz UI; (4) zbirni režim JESTE dostižan put za fail-open `ReportProsecnaCena` (Kooperanti/Vozači → globalna cena); (5) `ExcludeStornirano` celom aplikacijom ide kroz `FilterArray` — fail-open operator je infrastrukturno osetljiv.

### FM-0025 — `clsBlokIsplata.cls`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | Nema polja `Klasa` — multi-class redovi nerazlučivi | Visok | **Tačno** — cls:20-32 nema Klasa; identitet samo `otkupID` | P2 | Dodati `Klasa` + popuna u builderu | S |
| 2 | Nema version/stale podataka | Visok | **Tačno** — nema LoadedAt/verzije; CSV ne re-proverava (modBankaExportPregled.bas:369-374) | P2 | Pre `GenerisiNalogeCSV` sveže re-učitati saldo po OtkupID | M |
| 3 | Izvedena polja nezavisno mutable — kontradikcije moguće | Visok | **Tačno** — sva polja `Public`; invariant niko ne sprovodi | P2 | `Validate()` helper (Otvoren=Ukupan−Isplaćeno; Isplatiti≤Otvoren) pre exporta | S |
| 4 | Avans pool dupliran po redu — batch false-count | Visok | **Tačno** — isti `avansCache` saldo u svaki objekat (modBankaExportPregled.bas:131-136) | P2 | Shared snapshot oznaka; batch troši pool jednom po kooperantu | M |
| 5 | Nema export lifecycle/correlation | Visok | **Delimično** — polja nema, ali dizajn knjiži isplatu tek uvozom izvoda; ostaje rizik duplog CSV-a | P2 | Evidencija generisanih naloga (batch ID + datum) | L |
| 6 | Naziv „blok" skriva one-row (OtkupID) semantiku | Srednji | **Tačno** — grain je jedan OtkupID | P3 | Dokumentovati grain u komentaru klase | S |
| 7 | Nema root/generation ID | Srednji | **Tačno** — koncept ne postoji u modelu | Prihvaćeno | Ništa dok se model globalno ne uvede | — |
| 8 | Nema `IsManualOverride` — draft state podeljen | Srednji | **Tačno** — forma drži zaseban dictionary | P3 | Boolean flag umesto paralelnog dict-a | S |
| 9 | Nema validation/factory — parcijalan objekat moguć | Srednji | **Tačno** — default objekat prolazi | P3 | Isti `Validate()` kao #3 | S |
| 10 | Snapshot naziv/računa može zastareti | Srednji | **Tačno** — single-writer ublažava; duga modeless sesija rizik | P2 | Pre CSV re-lookup tekućeg računa | S |
| 11 | Nema valute/rounding pravila | Srednji | **Dizajnersko ograničenje** — sistem RSD + Double s tolerancijom 0,01 | Prihvaćeno | Ništa | — |
| 12 | Nema naziva stanice — dodatni lookup | Nizak | **Tačno** | P3 | Polje tek kad prikaz zatreba | S |
| 13 | Javna polja bez properties | Nizak | **Tačno** — standardan VBA DTO obrazac | Prihvaćeno | Ništa | — |

Bilans: 13 — 11 Tačno, 1 Delimično, 1 Dizajnersko; 6×P2, 4×P3, 3×Prihvaćeno.

### FM-0026 — `clsFilterParam.cls`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | Kolona samo fizički indeks — schema drift filtrira pogrešno polje | Visok | **Delimično** — polje jeste Long (cls:16), ali SVI pozivi rešavaju indeks iz imena runtime (`RequireColumnIndex`) — drift se sam prilagođava | P3 | Opciono `ColumnName` uz indeks za dijagnostiku | S |
| 2 | Operator nevalidiran — typo tiho protumačen | Visok | **Tačno** — `Init` bez Trim/UCase/enum (cls:21-29); uz fail-open Case Else typo tiho briše filter | P2 | `UCase$(Trim$(op))` + `Err.Raise` za nepodržan operator | S |
| 3 | BETWEEN ne zahteva `Value2` | Visok | **Tačno** — `Optional val2 = Empty`; `CDbl(Empty)=0` daje tih pogrešan interval | P2 | BETWEEN bez val2 → `Err.Raise` | S |
| 4 | Default objekat nevalidan ali prihvatljiv | Srednji | **Tačno** — `New` bez `Init` prolazi | P3 | Pokriva se #2 (prazan → greška) | S |
| 5 | Sva polja mutable posle dodavanja u kolekciju | Srednji | **Tačno** — reference semantika | P3 | Ništa hitno | — |
| 6 | Ponovni `Init` mutira istu instancu | Srednji | **Tačno** — vraća `Me` | P3 | Dokumentovati; pozivi su svuda `New`+`Init` | — |
| 7 | Variant coercion nedefinisan | Srednji | **Tačno** — semantika u `MatchesFilter` | P3 | Dokumentovati ugovor | S |
| 8 | `LIKE` semantika nedeklarisana | Srednji | **Tačno** — substring CI (modArrayUtils.bas:93) | P3 | Komentar: „LIKE = contains, case-insensitive" | S |
| 9 | Nema expected column type | Srednji | **Tačno** | Prihvaćeno | Ništa | — |
| 10 | Nema AND/OR grupa | Srednji | **Dizajnersko ograničenje** — svi consumeri homogeni AND | Prihvaćeno | Ništa | — |
| 11 | Nema null/empty operatora | Srednji | **Tačno** — `= ""` pokriva prazno | P3 | Ništa dok ne zatreba | — |
| 12 | Nema diagnostic `Describe` | Nizak | **Tačno** | P3 | Opciono | S |
| 13 | Nema Clone/freeze | Nizak | **Tačno** | Prihvaćeno | Ništa | — |

Bilans: 13 — 11 Tačno, 1 Delimično, 1 Dizajnersko; 2×P2, 8×P3, 3×Prihvaćeno.

### FM-0027 — `modArrayUtils.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | Unknown operator → `True` (fail-open) | Kritičan | **Tačno** — :94-95; kroz `ExcludeStornirano` (modHelpers.bas:179-182) nosi ceo storno sloj | P2 | `Case Else` → `Err.Raise "Unknown filter operator"` | S |
| 2 | Pretpostavlja 1-based 2D niz | Visok | **Delimično** — tačno (:20-21,31), ali svi izvori su `ListObject.Value` (1-based) | P3 | Assert `LBound=1` sa porukom | S |
| 3 | Nema validacije kolona — pad bez konteksta | Visok | **Delimično** — indeksi dolaze iz `RequireColumnIndex`; pad je glasan | P3 | Guard 1..UBound + opisna greška | S |
| 4 | Null/Error ćelija obara transformaciju | Visok | **Delimično** — pad je glasan exception, ne tihi rezultat | P3 | `IsError`/`IsNull` → no-match | S |
| 5 | BETWEEN ne validira granice | Visok | **Tačno** — :82-89; `CDbl(Empty)=0` tiho | P2 | Validirati Value1/Value2 jednom pre petlje | S |
| 6 | Group ključ raw `CStr` | Visok | **Tačno kao kod** (:246), ali `GroupBySum` NEMA callera (mrtav kod) | P3 | Obrisati `GroupBySum` | S |
| 7 | Nenumeričke sume tiho preskočene | Visok | **Tačno kao kod** (:225,:255), ali bez callera; obrazac živi u modIzvestaj (→ FM-0028 #21) | P3 | Obrisati mrtve funkcije | S |
| 8 | `=` case-sensitive, LIKE/sort nisu | Srednji | **Tačno** — :69 vs :93 vs :193 | P3 | Dokumentovati; storno vrednosti upisuje kod | S |
| 9 | `LIKE` je substring, ne wildcard | Srednji | **Tačno** — :93 `InStr` | P3 | Komentar/alias `CONTAINS` | S |
| 10 | Sort nije stabilan | Srednji | **Tačno** — QuickSort :136-170 | P3 | Tie-breaker originalni indeks | S |
| 11 | Comparator se menja po paru tipova | Srednji | **Tačno** — :176-194 | Prihvaćeno | Ništa — kolone homogene | — |
| 12 | Sekundarni sort bez posebnog smera | Srednji | **Tačno** — `ascending` zajednički | Prihvaćeno | Ništa | — |
| 13 | Empty postaje 0 u numeričkom filteru | Srednji | **Tačno** — `CDbl(Empty)=0` (:74-88) | P3 | `IsEmpty` guard → no-match | S |
| 14 | Group dictionary case-sensitive | Srednji | **Tačno** — :238; mrtav kod | Prihvaćeno | Ništa | — |
| 15 | `sumCols()` nevalidiran | Srednji | **Tačno** — :243; mrtav kod | Prihvaćeno | Ništa | — |
| 16 | Nema invalid/skipped count-a | Srednji | **Tačno** | P3 | Videti FM-0028 #21 (živi kod) | M |
| 17 | Double sume bez centralnog rounding-a | Srednji | **Tačno** | Prihvaćeno | Konzistentno (tolerancija 0,01) | — |
| 18 | `rowCount` je zapravo UBound | Nizak | **Tačno** — :20 | Prihvaćeno | Ništa | — |
| 19 | `colCount` je samo UBound | Nizak | **Tačno** — :21 | Prihvaćeno | Ništa | — |

Bilans: 19 — 16 Tačno, 3 Delimično; 2×P2, 10×P3, 7×Prihvaćeno. #6/#7/#14/#15 precenjeni: mrtav kod.

### FM-0028 — `modIzvestaj.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | Kartice počinju od nule — „saldo" je neto promena perioda | Kritičan | **Tačno** — runSaldo od 0 samo nad period-redovima (:636-648); amb kartica isto (:684) | P1 | Red „Početno stanje" pre `datumOd` (isti loop bez donje granice) | M |
| 2 | `ReportSaldoOM` meša periode — ambalaža all-time | Kritičan | **Delimično** — tačno (:205-231 vs :104-105), ali svesno („aktivni saldo" :271); problem je neoznačenost u UI | P2 | Header „Amb. (trenutno stanje)" u frmIzvestaj | S |
| 3 | Isplate kooperanta nisu station-attributed | Kritičan | **Tačno** — :116-121 sabira SVE isplate kooperanta u periodu, bez OMID/OtkupID | P1 | Filtrirati Novac po `COL_NOV_OM_ID` = stanica (avans-grana :139 to već radi) | M |
| 4 | `ReportManjak` nema OM filter | Kritičan | **Delimično** — tačno (:2162-2170), ali UI ne nudi Manjak pojedinačnom OM-u — nedostižno | P2 | OM grana ili Err za nepodržan tip | S |
| 5 | Bez prijemnice: 0% (RobaOM) vs 100% (Manjak) | Kritičan | **Tačno** — :1558-1577 vs :2257-2262 | P1 | U RobaOM „nema prijema" oznaka umesto 0 | S |
| 6 | Kupac per-vrsta uplata = prva stavka fakture | Kritičan | **Tačno** — prva stavka pobedi (modNovac.bas:783-789), cela uplata na tu vrstu (:427-438) | P1 | Raspodeliti uplatu srazmerno stavkama fakture | M |
| 7 | Ambalažni group key bez DokTip | Kritičan | **Delimično** — ključ jeste `DokID\|Tip` (:1933), ali ID-jevi su prefiksovani po tabeli — kolizija praktično nemoguća | P3 | Dodati dokTip u ključ | S |
| 8 | `OutputToSheet` ne čisti stari report | Kritičan | **Delimično** — funkcija ne čisti (:2695-2717), ali jedini caller čisti ceo sheet (modPrint.bas:42) | P3 | Clear footprint u samoj funkciji (za buduće callere) | S |
| 9 | Money-only kooperant → trenutna master stanica | Visok | **Tačno** — :107-114 master lookup umesto istorijske | P2 | `COL_NOV_OM_ID` reda umesto master lookup-a | S |
| 10 | Neraspoređena agrohemija globalna — ponavlja se po stanici | Visok | **Tačno** — :195-197 bez stanica provere; ulazi u UKUPNO svake stanice | P2 | Filtrirati blank-koop izlaze po stanici ili izuzeti iz UKUPNO | S |
| 11 | OM avans = periodski neto, ne raspoloživi saldo | Visok | **Tačno** — :138-155 samo period; label sugeriše stanje | P2 | Label „OM AVANS (promet perioda)" | S |
| 12 | `ReportAmbalaza` unknown tip → globalni report | Visok | **Tačno** — :1765-1795 bez Else; UI ne šalje druge tipove | P2 | `Else` → Empty/`Err.Raise` | S |
| 13 | `ReportProsecnaCena` unknown tip → OM/global grana | Visok | **Tačno** — :2030/:2064; DOSTIŽNO iz zbirnog moda za Kooperante/Vozače | P1 | Eksplicitni `Select Case`; Else → Empty (+ UI matrica, FM-0029 #3) | S |
| 14 | `ReportManjak` unknown tip → globalni report | Visok | **Tačno** — :2162-2170; dostižno zbirni Kooperanti | P2 | Isti fix kao #4/#13 | S |
| 15 | Nepoznat smer ambalaže postaje Izlaz | Visok | **Tačno** — :1858-1862, :1944-1950 | P2 | Case Ulaz/Izlaz/Else → data-quality greška | S |
| 16 | Kupac `Cena` = poslednja, ne prosečna | Visok | **Tačno** — :1141 overwrite; komentar „letzte" (:1233) | P2 | Ponderisana ili header „Poslednja cena" | S |
| 17 | Kupac ambalaža nije active saldo — ista kolona, druga semantika | Visok | **Tačno** — :1145 periodski zbir vs SaldoOM all-time | P2 | Header „Amb. (period)" | S |
| 18 | Detail i UKUPNO kupca — različiti algoritmi | Visok | **Tačno** — :1155 `GetUplataByVrsta` (samo >0) vs :1163-1190 direktan sken | P2 | UKUPNO = zbir per-vrsta redova + kontrolni red | S |
| 19 | `BrojZbirne` join bez generation scope-a | Visok | **Tačno** — :2211/:2227; prijemnice bez date filtera (:2224-2234) | P2 | Join preko `ZbirnaID` gde postoji; ograničiti prijemnice periodom | M |
| 20 | „Nema podataka" ne čisti stare ćelije | Visok | **Delimično** — :2696; ista mitigacija kao #8 | P3 | Uz #8 | S |
| 21 | Nevalidni source redovi tiho preskočeni | Visok | **Tačno** — obrazac svuda (:103, :361, :429, :493) | P2 | Brojati preskočene + upozorenje u statusu | M |
| 22 | Zbirni ambalažni report nema obećani UKUPNO | Srednji | **Tačno** — :1871-1884 vs komentar :1739 | P3 | Dodati UKUPNO red | S |
| 23 | Dictionary rezultati nesortirani | Srednji | **Tačno** — SaldoOM/Kupci/Isplata/PC/Manjak/Zbirni bez `SortArray` | P3 | Sortirati pre povratka | S |
| 24 | Running saldo tie order nestabilan | Srednji | **Tačno** — :628 QuickSort; krajnji saldo isti | P3 | Ref-ključ kao treći tie-breaker | S |
| 25 | Ref ključevi NOV/MAG/AMB bez ID reda | Srednji | **Tačno** — :460/:517/:593 vs „OTK\|id" :403 | P3 | ID kad drill-down zatreba | M |
| 26 | Raw tekst ključevi fragmentiraju grupe | Srednji | **Tačno** — :54, :2097 bez Trim | P3 | `Trim$` na dictionary ključevima | S |
| 27 | `\|` delimiter bez escape-a | Srednji | **Tačno** — :2365/:2391; sistemski ID-jevi ga ne sadrže | Prihvaćeno | Ništa | — |
| 28 | Period nevalidiran (obrnut = prazan report) | Srednji | **Tačno** — nigde `datumOd<=datumDo` | P3 | Provera u frm `btnUnos` | S |
| 29 | DateTime `datumDo` može izostaviti kraj dana | Srednji | **Nije proverivo statički** — zavisi da li podaci nose vreme | P3 | `Int()` na cell datum ili `datumDo+1` ekskluzivno | S |
| 30 | PDF Sub bez success rezultata | Srednji | **Tačno** — :993-1038; `OFF`/Nothing → tiho ništa | P3 | Function sa putanjom/False + poruka u formi | S |
| 31 | PDF prepisuje isti period | Srednji | **Tačno** — :1021-1022 deterministična putanja | Prihvaćeno | Ništa — poželjno | — |
| 32 | Cell-by-cell output | Srednji | **Tačno** — :2712-2717 | P3 | `Resize(r,c).Value = data` | S |
| 33 | Double bez centralnog rounding-a | Srednji | **Tačno** | Prihvaćeno | Konzistentno | — |
| 34 | Enum ne pokriva sve report funkcije | Nizak | **Tačno** — :11-19; enum se NIGDE ne koristi | P3 | Obrisati mrtav enum | S |
| 35 | Helperi fail-soft gube nazive | Nizak | **Tačno** — `BuildOtkupBrojDokDict` EH vraća delimičan dict (:839-841) | P3 | Prihvatljiva degradacija; LogErr opciono | — |

Bilans: 35 — 29 Tačno, 5 Delimično, 1 Nije proverivo; 5×P1, 13×P2, 14×P3, 3×Prihvaćeno.

### FM-0029 — `frmIzvestaj.frm`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | Stari podaci + novi period u statusu/štampi | Kritičan | **Tačno** — status/naslov iz `txtDatum*` (:603-604, :1446-1447), podaci iz `m_cur*` (:662-666); nema change handlera | P1 | Status/štampu graditi iz `m_curOd/m_curDo`; na izmenu datuma „nije osveženo" | S |
| 2 | Lazy report greška potpuno tiha | Kritičan | **Tačno** — `CleanFail` samo LogErr (:702-705); stara lista + zelen status ostaju | P1 | Clear aktivne liste + status „Greška" + MsgBox | S |
| 3 | Zbirni mode nudi nevalidne kombinacije | Kritičan | **Tačno** — :488-492 tabovi 5/6/7 SVIM tipovima; Kooperant→Empty, PC/Manjak→globalno pod pogrešnim kontekstom | P1 | Tabove 5/6/7 samo za OM/Kupci/Vozaci (validna matrica) | S |
| 4 | Revers spaja više tipova ambalaže | Kritičan | **Tačno** — `tipAmb` sa prvog reda (:2116), `kolAmb` sabira sve redove bez tipa (:2121,:2126) | P1 | Preneti `TipAmbalaze` izabranog reda u match kriterijum | S |
| 5 | Revers uključuje stornirane redove | Kritičan | **Tačno** — `GetTableData(TBL_AMBALAZA)` bez `ExcludeStornirano` (:2089) | P1 | `d = ExcludeStornirano(d, TBL_AMBALAZA)` | S |
| 6 | `m_SetupDone=True` pre uspešnog setup-a | Visok | **Tačno** — :90-91; EH log-only → parcijalna forma bez retry | P2 | True tek na kraju; u EH reset + poruka | S |
| 7 | `m_IsInitializing` može trajno ostati True | Visok | **Tačno** — :103/:193; EH ne resetuje → AutoRefresh/toggle mrtvi | P2 | U EH: `m_IsInitializing=False` | S |
| 8 | Kupac/stanica resolve samo po nazivu | Visok | **Tačno** — `LookupValue` Naziv→ID (:646-648); dupli naziv → prvi | P2 | Combo „Naziv (ID)" + `ExtractIDFromDisplay` | M |
| 9 | Prazan `entitetID` neblokiran → globalni scope | Visok | **Tačno** — :644-656 bez provere | P2 | `If entitetID = "" Then MsgBox + Exit` | S |
| 10 | Mrtav `cmbVrstaRobe` filter | Visok | **Tačno** — samo punjenje (:170-178), nula čitanja | P2 | Sakriti kontrolu ili implementirati filter | S |
| 11 | Status broji UKUPNO/summary redove | Visok | **Tačno** — `ListCount` (:603) | P3 | Oduzeti kontrolne redove | S |
| 12 | Modeless forma ne invalidira stale cache | Visok | **Tačno** — Activate posle setup-a odmah izlazi; `m_genTabs` preživi | P2 | Na Activate „podaci možda zastareli" ili invalidirati | M |
| 13 | Naslov „Saldo" skriva period-from-zero | Visok | **Tačno** — headeri :1496/:1507/:1555, štampa :1274-1397 | P2 | „Saldo (promet perioda)" dok se ne doda početno stanje | S |
| 14 | Pending prijemnica prikazana kao 0% manjka | Visok | **Tačno** — :901-905 formatira nule kao rezultat | P1 | Uz FM-0028 #5: „nema prijema" umesto „0 / 0.00%" | S |
| 15 | Deljeni detail state može ostati sa drugog taba | Visok | **Delimično** — `SetVisible` ne briše, ali svaki `Show*` resetuje ID polja + print guard; ostaje vizuelno stale | P3 | `KarticaDetalji_Clear` u `mpReports_Change` | S |
| 16 | Revers sabira duplikate/generacije | Visok | **Tačno** — :2111-2129 bez exact-pair/generation provere | P2 | Uz #5: posle storno filtera upozoriti na >2 noge | S |
| 17 | Nedostajući datum reversa postaje danas | Visok | **Tačno** — :2142 `datum = Date` | P2 | Greška „dokument bez datuma" umesto tihe zamene | S |
| 18 | Generic print koristi display stringove | Srednji | **Tačno** — :1430-1434 | P3 | Prihvatljivo; opciono raw model | — |
| 19 | Kartica PDF može biti drugog perioda od ekrana | Srednji | **Tačno** — :1238-1242 `txtDatum*` umesto `m_cur*` | P2 | Koristiti `m_curOd/m_curDo` (isti koren kao #1) | S |
| 20 | Runtime tab build bez retry | Srednji | **Tačno** — EH postavlja built=True (:1739-1741, :1854-1856) | P3 | Prihvatljivo (anti-dupli page) | S |
| 21 | LoadEntiteti fail-soft — prazna lista bez poruke | Srednji | **Tačno** — `On Error GoTo done` (:370,:402) | P3 | LogErr + status poruka | S |
| 22 | Stanice/kupci bez vidljivog ID-ja | Srednji | **Tačno** — :375-377 samo Naziv | P3 | Uz #8 | — |
| 23 | Drill-down greške nevidljive (OERN) | Srednji | **Tačno** — :1216, :1947, :1977, :1994 | P3 | LogErr umesto golog Resume Next | S |
| 24 | Zbirna ambalaža bez totala, UI ne signalizira | Srednji | **Tačno** — core ne vraća UKUPNO (FM-0028 #22) | P3 | Uz FM-0028 #22 | — |
| 25 | Generic print bez rezultata/potvrde | Srednji | **Tačno** — `PrintIzvestaj` Sub; export u OERN bloku (modPrint.bas:54-60) | P3 | Putanja + potvrda; skinuti OERN oko exporta | S |
| 26 | Redosled perioda nevalidiran | Srednji | **Tačno** — :618-620 samo `CDate` | P3 | `If datumOd > datumDo Then MsgBox` | S |
| 27 | Status zelen za bilo koju nepraznu listu | Srednji | **Tačno** — :599-605 bez freshness | P3 | Vezati za `m_genTabs` + `m_cur*` (uz #1) | — |
| 28 | `UpdateUnosButtonState` možda mrtav | Nizak | **Tačno** — definisan :758, nula poziva | P3 | Obrisati ili pozvati | S |
| 29 | `PrijemniceZaOtpremnicu` bez callera | Nizak | **Tačno** — definisan :2172, nula poziva | P3 | Obrisati (git čuva istoriju) | S |
| 30 | Source caption „UserForm1" | Nizak | **Tačno** — :3; runtime `FixFormCaptions` postavlja pravi | Prihvaćeno | Ništa | — |

Bilans: 30 — 29 Tačno, 1 Delimično; 6×P1, 10×P2, 13×P3, 1×Prihvaćeno.

**Bilans bloka F (110):** 96 Tačno / 11 Delimično / 2 Dizajnersko / 1 Nije proverivo; **11×P1**, 31×P2, 51×P3, 17×Prihvaćeno. Najisplativiji paket: `StampajReversAmbDok` (storno + tip ambalaže, 2×P1 za S), zbirna tab-matrica (P1, S), status/štampa iz `m_cur*` (2×P1, S), station-attribution isplata i „nema prijema" status (2×P1, S–M).

---

## Blok G — Kartice, štampa, ambalaža, faktura (FM-0030…FM-0034, 133 stavke)

### FM-0030 — `modKarticaDetalji.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | Current ID nije vezan za tab/list/row; cross-tab štampa | Kritičan | **Tačno** — module-state (:27-30); dugme čita getter (frmIzvestaj.frm:1960); promena taba ne čisti (frmIzvestaj.frm:737-749) | P1 | `KarticaDetalji_Clear` u `mpReports_Change`; dugme čita ref selektovanog reda | S |
| 2 | Deselekcija (`idx<0`) ne briše stari print target | Kritičan | **Tačno** — guard pre brisanja (:154-156, :227-237); ali regeneracija zove Clear — prozor uzak | P2 | Clear pre `Exit Sub` kod `idx<0` | S |
| 3 | Stornirani otkup ostaje print target | Visok | **Tačno** — bez `ExcludeStornirano`, prvi pogodak (:352-371) | P2 | `ExcludeStornirano` u `ShowOtkupDetails`; kod storna poruka | S |
| 4 | OtpremnicaID se prihvata bez validacije | Visok | **Tačno** — `mCurOtpremnicaID=otpID` odmah (:273); ID dolazi iz ref-ključa istog reda | P2 | Existence+storno provera pre postavljanja | S |
| 5 | Ambalažni target iz ref teksta bez ledger potvrde | Visok | **Tačno** — parse i set bez provere (:326-334) | P2 | Potvrditi DokID+DokTip u tblAmbalaza pre set-a | S |
| 6 | Mixed snapshot (stari red + svež lookup) | Visok | **Tačno** — kolone liste (:275-284) + svež `LookupValue` (:288-293) | P3 | Jedan izvor; ili label „trenutna cena" | S |
| 7 | Duplicate OtkupID uzima prvi red | Visok | **Delimično** — kod potvrđen (:361-366), ali OtkupID generiše `GetNextID` (duplikat = anomalija) | P3 | Brojati pogotke; >1 → upozorenje | S |
| 8 | `BrojZbirne` join bez generation scope-a | Visok | **Delimično** — join potvrđen (:301-320), isključuje stornirane; samo prikaz | P3 | Tehnički ključ (ZbirnaID) kada postoji | M |
| 9 | NOV/MAG nema row identity | Srednji | **Dizajnersko ograničenje** — ref bez ID (:16, :255-267) | P3 | Proširiti ref-ključ row ID-jem | M |
| 10 | Schema problem postaje prazno | Srednji | **Tačno** — `CellVal` vraća "" (:423-431) | P3 | Prikaz „n/a (kolona)" | S |
| 11 | Nenumeričko postaje 0 | Srednji | **Tačno** — `NumOf` (:433-435) | P3 | Prikaz sirove vrednosti uz oznaku | S |
| 12 | Master nazivi su trenutni | Srednji | **Dizajnersko ograničenje** — live lookup (:437-478); panel je pregled uživo | Prihvaćeno | Bez izmene | — |
| 13 | `AddPair` guta greške | Srednji | **Tačno** — OERN (:416-420) | P3 | Log; blokirati print kad panel nepotpun | S |
| 14 | Parcijalni build nema cleanup | Srednji | **Tačno** — EH ne briše naslov (:120-122); retry pada na duplo ime | P3 | U EH ukloniti kreirane kontrole | S |
| 15 | Singleton state (jedna forma) | Srednji | **Dizajnersko ograničenje** — jedna report forma u praksi | Prihvaćeno | Klasa-instanca tek uz drugu formu | M |
| 16 | Linearni scan/lookup po kliku | Srednji | **Tačno** — cela tblOtkup po kliku (:352); obim mali | P3 | Keš niza po report sesiji | M |

Bilans: 16 — 11 Tačno, 2 Delimično, 3 Dizajnersko/Prihvaćeno; 1×P1, 4×P2, 9×P3, 2×Prihvaćeno.

### FM-0031 — `modPrint.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | Multi-ID: nedostajući ID se tiho preskače | Kritičan | **Tačno** — petlja bez not-found greške (:387-429; :883-925; :1147-1175); realni calleri šalju sveže ID-jeve | P2 | Brojati nepronađene; >0 → poruka/prekid | S |
| 2 | Hibridni dokument: header sa prvog reda | Kritičan | **Tačno** — header samo kad `koopID=""` (:416-422; :912-918; :1163-1168), bez same-BrDok/partner provere | P2 | Validirati da svi redovi dele BrDok/kooperanta/datum | S |
| 3 | Reprint storniranog kroz fallback | Kritičan | **Tačno** — `If ids="" Then ids=otkupID` (:334) → `FillOtkupSablon` čita sirovu tabelu (:354); put postoji iz frmIzvestaj (:1965, :2067) | **P1** | Umesto fallback-a poruka „otkup je storniran"; ne štampati | S |
| 4 | Historical PDV drift (trenutna stopa pri reprintu) | Kritičan | **Tačno** — stopa iz configa pri štampi (:120-121, :382-383, :878-879); čuva se samo bruto cena | P2 | Snapshot stope/neto pri unosu; reprint iz snapshot-a | M |
| 5 | Revers hidden direction state (`m_izdAmbPrijem`) | Kritičan | **Delimično** — state postoji (:21, :1349, :1394-1396, :1463), ali JEDINI caller Export-a je Output koji ga prethodno postavi (:1360) — nema živog pogrešnog puta | P3 | `prijem` parametar u `ExportIzdavanjeAmbalazePDF` | S |
| 6 | Historical revers saldo drift | Kritičan | **Tačno** — `IzdAmbSaldoVal` čita tekući ledger (:1763-1782); reprint put postoji (frmIzvestaj.frm:2087-2164) | P2 | Istorijski saldo do DokumentID (obrazac `GetKooperantAmbOpening`) | M |
| 7 | Otpremnica koristi otkupni/PDV obrazac | Kritičan | **Delimično** — sadržaj potvrđen (:218-219; PDV :180-183; koop prazan :172), ali svesna reuse odluka (komentar :67-68) | P2 | Poseban otpremnica layout ako biznis traži | L |
| 8 | Reprint grupiše samo po `BrDok` | Kritičan | **Tačno** — bez koop/stanica/datum scope-a (:327-333) | P2 | Scope: BrDok + KooperantID + datum | S |
| 9 | Direct fill ne filtrira storno | Visok | **Tačno** — sirov `GetTableData` u sva 4 Fill-a (:354, :852, :1119, :110) | P2 | `ExcludeStornirano` u Fill funkcijama | S |
| 10 | Duplicate input ID duplira stavke/iznose | Visok | **Tačno** — Split bez dedupe (:376-428); nijedan caller ne šalje duplikat | P3 | Dedupe Dictionary pre petlje | S |
| 11 | Multi-class ambalaža: tip/izdato sa prvog reda | Visok | **Delimično** — kod potvrđen (:420-421), ali model garantuje isti tip po bloku i izdato na tačno jednom redu (modOtkup.bas:200/238, 220-227) | P3 | Sabrati KolAmbIzd preko svih redova (robusnost) | S |
| 12 | Output Sub-ovi bez result contract-a | Visok | **Delimično** — potvrđeno (:70, :268, :803, :1054, :1341), ali svestan best-effort post-commit dizajn; Export* vraćaju putanju | P3 | Status/putanja iz Output*; UI ne tvrdi uspeh | M |
| 13 | `PrintIzvestaj` potpuno guta PDF grešku | Visok | **Tačno** — OERN oko exporta, bez loga (:54-60) | P2 | Ukloniti OERN; `Dir$` provera + poruka | S |
| 14 | `_Print` ownership nije potvrđen | Visok | **Tačno** — preuzima postojeći sheet i `Cells.Clear` (:34-42) | P3 | Ownership marker pre Clear | S |
| 15 | Otpremnica PDF deterministički (overwrite) | Visok | **Tačno** — bez timestampa (:85) vs prijemnica (:1094) | P3 | Timestamp sufiks | S |
| 16 | Trenutna klauzula/master menjaju reprint | Visok | **Tačno** — klauzula/rok/seller/koop pri renderu (:432-481) | P2 | Deo document-snapshot inicijative (uz #4) | L |
| 17 | 0% stopa pada na default | Visok | **Tačno** — `stopa<=0 → default` (:121, :383, :879); 0% nerealna konfiguracija | P3 | Razlikovati prazno od eksplicitne nule | S |
| 18 | Otkup layout nije za >2 stavke | Visok | **Tačno** — bez guarda (:736-737, :763-764); >2 nastaje tek uz dupli BrDok | P3 | Guard `nStavke>2` → upozorenje | S |
| 19 | Faktura variable merge se ne unmerge-uje | Visok | **Tačno** — cleanup bez UnMerge (:1896-1900), total merge ostaje (:1923-1924); sledeća faktura sa VIŠE stavki renderuje kroz stari merge | **P1** | `.UnMerge` u cleanup opseg `FillFakturaSablon` | S |
| 20 | Faktura/spec total nije reconciled | Visok | **Tačno** — upis bez SUM provere (:1928, :2607-2610, :2752-2755) | P3 | Warning ako SUM(stavke)≠total | S |
| 21 | Template sheetovi čuvaju osetljive podatke | Visok | **Delimično** — potvrđeno (spec vidljiv :2509, :2518, :2661, :2670), ali workbook ionako sadrži iste podatke | P3 | Očistiti/sakriti spec sheet posle exporta | S |
| 22 | Filename sanitizacija nepotpuna | Srednji | **Tačno** — samo `/`, `" + "`; bonus: `"\\"` menja dupli, ne jednostruki backslash (:83, :293) | P3 | Centralni `SanitizeFileName` | S |
| 23 | Otkup filename koristi ID, ne BrDok | Srednji | **Tačno** — komentar kaže brDok (:286), kod ID (:293-294) | P3 | BrDok + timestamp | S |
| 24 | `PrintPrijemnica` zaobilazi mode | Srednji | **Tačno** — direktan `PrintOut` (:1073-1076); poziva frmIzvestaj:2065 | P3 | Preusmeriti na `OutputPrijemnica` | S |
| 25 | Fiksne cleanup granice (80/300/1000) | Srednji | **Tačno** — :1896, :2055, :2250+:2254, :2592, :2737 | P3 | Guard `nRows>granica` | S |
| 26 | Version marker nije schema check | Srednji | **Tačno** — jedna ćelija (:1823, :2152, :2348, :2509, :2661) | P3 | Prihvatljivo; opciono named-range | S |
| 27 | Neki template-i nisu verzionisani | Srednji | **Delimično** — Ensure bez verzije, ali Fill svaki put briše i ponovo crta ceo sheet — stale layout nema efekat | P3 | Bez izmene | M |
| 28 | Application state se ne vraća dosledno | Srednji | **Tačno** — EH nekih Fill-ova `=True` umesto `oldScreen` (:2133, :2330, :2490, :2643, :2788) | P3 | U EH vraćati sačuvano stanje | S |
| 29 | PageSetup failure je tih | Srednji | **Dizajnersko ograničenje** — nameran hardening bez štampača (:1279-1293) | P3 | Log pri grešci | S |
| 30 | Sledljivost bez input validacije | Srednji | **Tačno** — `CDate/CDbl` bez guarda (:2258-2274, :2329-2331) | P3 | Shape/datum provere s porukom | S |
| 31 | Negativan manjak bez statusa | Srednji | **Tačno** — sirov prikaz (:2291-2300) | P3 | Label „Višak" kad je negativan | S |
| 32 | Kartice formalizuju saldo od nule | Srednji | **Tačno** — šablon nema red početnog stanja (:2050-2109); nasleđeno iz modIzvestaj | P3 | Red „Početno stanje" kad ga modIzvestaj da | M |
| 33 | Nema output audit/hash | Srednji | **Tačno** — svesno odsustvo | P3 | Tek uz `DocumentRenderResult` inicijativu | L |
| 34 | Dupliran `Attribute VB_Name` komentar | Nizak | **Tačno** — linije 1-2; artefakt alata | Prihvaćeno | Ništa | — |

Bilans: 34 — 28 Tačno, 6 Delimično; 2×P1 (#3 reprint storniranog, #19 faktura merge), 9×P2, 22×P3, 1×Prihvaćeno.

### FM-0032 — `modDocStyle.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | `DocPrintWs` default-to-print (sve osim PREVIEW štampa) | Kritičan | **Tačno** — else-grana `PrintOut` (:210-216); ali SVI pozivi gate-ovani na PRINT/PREVIEW — nema živog pogrešnog puta | P2 | Strict `Select Case`; ostalo `Err.Raise` | S |
| 2 | PageSetup fail-silent (perforacija bez signala) | Kritičan | **Tačno** — OERN preko cele procedure (:172-191); nameran hardening | P2 | Postcondition provera Zoom/PrintArea + log | S |
| 3 | Invalid mode pada na default | Visok | **Tačno** — `Case Else → defMode` (:199-207) | P3 | Log za neprazan nepoznat mode | S |
| 4 | `defMode` se ne validira | Visok | **Tačno** — vraća `UCase(defMode)` (:205); interni pozivi šalju literale | P3 | Validirati protiv skupa | S |
| 5 | Legal config fail-open (klauzula) | Visok | **Tačno** — `DocConfigOr` guta grešku (:25-31), pad na hardkod (:35-43) | P3 | Razlikovati grešku čitanja od praznog; log | S |
| 6 | PDF nema result contract | Visok | **Tačno** — Sub bez provere fajla (:157-162); greška se propagira (nema OERN) | P3 | Function + `Dir$`/`FileLen` provera | S |
| 7 | Istorijski seller/legal drift | Visok | **Tačno** — `GetConfigValue` pri renderu (:101-110) | P2 | Deo document-snapshot inicijative | L |
| 8 | `PrintCommunication` state se ne čuva | Visok | **Tačno** — False→True bezuslovno (:174, :189) | P3 | Sačuvati/vratiti prethodno stanje | S |
| 9 | PageSetup nema postcondition | Visok | **Tačno** — isto mesto kao #2 (:175-188) | P2 | Isto kao #2 | S |
| 10 | Logo potpuno fail-silent | Srednji | **Tačno** — `On Error GoTo done` prazan (:61-78); logo opcion | P3 | Log jednom po sesiji | S |
| 11 | Logo se deformiše (52×40 bez aspect) | Srednji | **Tačno** — fiksne dimenzije (:66-76) | P3 | `LockAspectRatio`/scale-to-fit | S |
| 12 | Logo se može duplirati | Srednji | **Delimično** — helper ne briše, ali calleri brišu Shapes pre fill-a | P3 | Stabilno ime shape-a + replace | S |
| 13 | Windows path separator | Srednji | **Dizajnersko ograničenje** — aplikacija je Windows-only | Prihvaćeno | Ništa | — |
| 14 | Seller polja nisu validirana | Srednji | **Tačno** — bez provera (:100-113) | P3 | Warning za prazan SELLER_NAME/PIB | S |
| 15 | Title block nije idempotentan | Srednji | **Tačno** — Merge bez UnMerge (:128, :136); calleri zovu na svežem sheetu | P3 | UnMerge ciljanih redova pre Merge | S |
| 16 | PDF nije atomaran | Srednji | **Tačno** — direktan upis na finalnu putanju (:159-161) | P3 | Temp+rename tek uz result contract | M |
| 17 | Folder/visibility su caller odgovornost | Srednji | **Dizajnersko ograničenje** — podela odgovornosti | Prihvaćeno | Dokumentovati ugovor | S |
| 18 | Nema filename sanitizacije u shared sloju | Srednji | **Tačno** — helper ne postoji | P3 | `DocSanitizeFileName`; koristi modPrint | S |
| 19 | Nema template ownership/version helpera | Srednji | **Tačno** — ne postoji | P3 | Zajednički Ensure helper | M |
| 20 | `lastRow` se ne validira | Srednji | **Tačno** — sirov u PrintArea (:187) | P3 | Guard `lastRow>=1` | S |
| 21 | Nema output audit-a | Srednji | **Tačno** — ne postoji | P3 | Deo buduće telemetrije | L |
| 22 | Boje nisu konfigurabilne | Nizak | **Tačno** — hardkod RGB (:12-22) | Prihvaćeno | Ništa | — |
| 23 | `Ziro` nije lokalizovan | Nizak | **Tačno** — hardkod labela (:108-110) | Prihvaćeno | Ništa | — |
| 24 | Rich-text greška je tiha | Nizak | **Tačno** — OERN oko `Characters` (:89-91); dekorativno | P3 | Prihvatljivo | — |
| 25 | Dupliran `Attribute VB_Name` komentar | Nizak | **Tačno** — linije 1-2 | Prihvaćeno | Ništa | — |

Bilans: 25 — 21 Tačno, 1 Delimično, 3 Dizajnersko/Prihv.; 0×P1, 4×P2, 16×P3, 5×Prihvaćeno.

### FM-0033 — `modAmbalaza.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | Vozački znak suprotan invariantu (otvorena otpremnica negativna) | Kritičan | **Tačno** — formula `saldo=Izlaz−Ulaz` (:584) protivreči SOPSTVENIM komentarima koda (:58-59 „otvorena otpremnica = pozitivan saldo"; :512-513) — interna kontradikcija, ne poređenje sa starim draftom. Jedini potrošač je smoke-test → nema produkcionog prikaza | P2 | Okrenuti na `Ulaz−Izlaz` pre prvog UI potrošača; test na znak | S |
| 2 | Physical-row historical cutoff | Kritičan | **Tačno** — granica po fizičkom redu (:366-387); append-only pretpostavka dokumentovana (:307-312); korisnički sort tabele je krši | P2 | Cutoff po AmbID sekvenci | M |
| 3 | Missing-block fallback = ceo trenutni saldo | Kritičan | **Tačno** — `minIdx=0 → cutoff=kraj` (:381-387); nameran legacy fallback, ali tiho pogrešan istorijat na dokumentu | P2 | Fallback označiti (log/oznaka na dokumentu) | S |
| 4 | Schema guard ne štiti positional insert | Kritičan | **Tačno** — `Array→AppendRow` (:157-170); guard proverava postojanje, ne redosled (:113-126) | P2 | Upis po imenu ili order-assert | M |
| 5 | Read greška izgleda kao Empty/0 (false-zero) | Kritičan | **Tačno** — EH→Empty (:299-301), →0 (:407-409, :442-444); nula ulazi u dokumente (modPrint.bas:174, :461, :963) | P2 | Razdvojiti „nema pokreta" od greške | M |
| 6 | Nema idempotency/pair invariant | Kritičan | **Delimično** — odsustvo potvrđeno, ali pozivi su unutar TX wrappera; rizik samo van-TX retry — single-writer | P2 | Idempotency ključ uz budući transfer model | L |
| 7 | Public non-TX writer (ugovor u komentaru) | Visok | **Tačno** — :17-20 | P3 | Naming/`RequireTxContext` konvencija | M |
| 8 | Entitet/dokument se ne validiraju | Visok | **Tačno** — nema master provera (:77-111); calleri šalju ID iz comba | P3 | Opciona existence provera po tipu | M |
| 9 | DokumentID/Tip nisu obavezni | Visok | **Tačno** — Optional "" (:135-137) | P3 | Obavezno za poslovne tokove; izuzetak migracija | S |
| 10 | Multi-user AmbID kolizija | Visok | **Delimično** — `GetNextID` max+1 (:150); single-writer | P2 | Retry/unique tek uz multi-user | M |
| 11 | Historical stanica saldo je current-only | Visok | **Tačno** — bez cutoff parametra (:420-445); reprint drift kroz modPrint (:174, :963) | P2 | `GetStanicaAmbOpening` sa DokumentID granicom | M |
| 12 | Vozački period je promet od nule | Visok | **Tačno** — samo pokreti perioda (:516-524); trenutno bez potrošača | P3 | Dokumentovati; opening uz prvog potrošača | M |
| 13 | Nepoznat EntitetTip se prihvata | Visok | **Tačno** — samo neprazan (:107-110) | P3 | Enum validacija tri tipa | S |
| 14 | Legacy negativne/decimalne količine na read | Visok | **Tačno** — `IsNumeric`+`CLng` (:259-271, :395-399) | P3 | Read guard: celobrojno ≥0 | S |
| 15 | Corrupt EntitetID/Tip tiho nestaje | Visok | **Tačno** — `AmbText`→"" pre matchinga (:30-36, :248-249) | P3 | Brojati/logovati redove praznog entiteta | S |
| 16 | Master fallback može upisati nedozvoljen tip | Visok | **Tačno** — fallback 12/1, 6/1 (:602-617); nameran backward-compat | P3 | UI warning kad je fallback aktivan | S |
| 17 | DokumentTip boundary nije proveravan | Srednji | **Delimično** — kod potvrđen (:373), ali ID prefiksi (OTK-/PRJ-/…) koliziju čine praktično nemogućom | P3 | DokumentTip u match | S |
| 18 | Opening/current različita strictness | Srednji | **Tačno** — opening preskače (:395-399), current diže grešku (:254-276) | P3 | Ujednačiti ugovore | S |
| 19 | DateTime poslednjeg dana | Srednji | **Tačno** — `d > datumDo` (:523); datumi iz formi su date-only | P3 | `Int(d)` poređenje | S |
| 20 | Invalid datumi se preskaču | Srednji | **Tačno** — `GoTo NextRow` (:517) | P3 | Brojati preskočene | S |
| 21 | Dictionary case fragmentacija tipova | Srednji | **Delimično** — binary compare (:243-244, :504-505), ali tipovi dolaze iz master comba | P3 | `vbTextCompare` na saldo dict | S |
| 22 | Prvi culture match je autoritativan | Srednji | **Tačno** — prvi neprazan hit (:636-647) | P3 | Preferirati exact vrsta+sorta pogodak | S |
| 23 | Nema stable TipAmbalazeID | Srednji | **Dizajnersko ograničenje** — naziv je business key kroz ceo sistem | P3 | Samo uz veliku migraciju | L |
| 24 | Nema transfer/counterparty ID | Srednji | **Dizajnersko ograničenje** — 10-kolonska šema; noge vezuje DokumentID konvencija | P3 | `AmbTransferID`/`LegRole` uz migraciju šeme | L |
| 25 | Nema domain-level success audita | Srednji | **Tačno** — samo `LogErr`; DataAccess audit kolone postoje | P3 | Prihvatljivo | — |
| 26 | `Long` konverzija i overflow | Srednji | **Tačno** — `As Long` (:132-133); gajbice male vrednosti | P3 | Prihvatljivo; opciono guard | — |
| 27 | Output tipovi nisu sortirani | Nizak | **Tačno** — dict redosled (:285-294) | Prihvaćeno | `SortArray` ako UI zatraži | S |
| 28 | Dupliran `Attribute VB_Name` komentar | Nizak | **Tačno** — linije 1-3 | Prihvaćeno | Ništa | — |

Bilans: 28 — 22 Tačno, 4 Delimično, 2 Dizajnersko; 0×P1, 8×P2, 18×P3, 2×Prihvaćeno. Napomena uz #1: nalaz stoji i protiv TEKUĆEG koda (formula :584 vs komentari :58-59/:512-513); P1 izbegnut jedino jer funkcija još nema produkcionog potrošača.

### FM-0034 — `modFaktura.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | Kupac prijemnice se ne proverava (cross-buyer) | Kritičan | **Delimično** — u modulu potvrđeno (validaciona petlja :184-241 ne čita Prijemnica.KupacID), ali jedini UI caller lista SAMO prijemnice tog kupca (frmFakturisanje.frm:301) i čisti listu pri promeni kupca | P2 | U petlji porediti `Prijemnica.KupacID = kupacID` (defense-in-depth) | S |
| 2 | Duplicate PrijemnicaID u tabeli (rows(1)) | Kritičan | **Tačno** — bez `Count=1` guarda (:188-206); ispravan obrazac postoji u `RequireSingleFakturaRow` (:631-635) | P2 | `If rows.Count > 1 Then Err.Raise` | S |
| 3 | Public non-TX base `CreateFaktura` | Kritičan | **Tačno** — Public (:94); svi produkcioni pozivi kroz `_TX` | P2 | Učiniti Private | S |
| 4 | Multi-user double-invoice race | Kritičan | **Delimično** — odsustvo CAS/claim potvrđeno; single-writer | P2 | Re-check availability tik pre write-a | M |
| 5 | Positional header (21) / stavka (9) insert | Kritičan | **Tačno** — Array :260-282 i :304-314; guard samo imenovani podskup (:114-136) | P2 | Upis po imenu ili order-assert | M |
| 6 | Status i datum nisu atomarni | Kritičan | **Delimično** — dva `RequireUpdateCell` bez TX (:565-589), ALI procedura je idempotentna: sledeći poziv sanira raskorak | P3 | TX omot ili dokumentovati samoizlečenje | S |
| 7 | Kupac ne mora postojati | Visok | **Tačno** — samo neprazan (:98-101); UI bira iz comba | P3 | Existence provera kupca | S |
| 8 | Receipt lifecycle nedovoljno validiran | Visok | **Delimično** — storno/fakturisano/FakturaID pokriveni (:662-664); nabrojana stanja ne postoje u modelu | P3 | Bez izmene dok model ne dobije ta stanja | — |
| 9 | FakturaID/broj concurrency (max+1) | Visok | **Delimično** — `GetNextID` (:139), `GenerateBrojFakture` (:353-397); single-writer | P2 | Unique provera posle generisanja | M |
| 10 | Wrapper guta originalnu grešku (vraća "") | Visok | **Tačno** — EH→"" bez rethrow (:49-92); UI ćuti/opšta poruka | P3 | ByRef errMsg ili rethrow business grešaka | S |
| 11 | Rollback failure nije vidljiv | Visok | **Tačno** — `RollbackTx` pod OERN (:58, :83); neuspeo rollback = tiha parcijalna faktura | P2 | Kritična poruka „zatvori bez snimanja" pri padu rollback-a | S |
| 12 | Case-sensitive `"Da"` u availability | Visok | **Tačno** — exact (:662-664) vs `UCase="DA"` u PrintFaktura (:435); modul sam piše kanonsko „Da" | P3 | UCase normalizacija u helperu | S |
| 13 | Print detail/header nisu reconciled | Visok | **Tačno** — ukupno iz headera (:486-487), stavke zasebno (:467-479); oba iz istog create-a | P3 | Warning ako SUM≠header (±0.01) | S |
| 14 | Nenumerička print stavka postaje 0 | Visok | **Tačno** — fallback 0/0 (:470-472) | P3 | `Err.Raise` pri nenumeričkom snapshotu | S |
| 15 | Buyer legal snapshot ne postoji | Visok | **Tačno** — čuva se samo kupacID (:264); print čita trenutni naziv (:444) | P3 | Snapshot naziva/PIB-a u header | M |
| 16 | Datum plaćanja nije datum uplate | Visok | **Tačno** — `Date` pri recompute (:573-574) | P3 | Izvesti iz zatvarajuće uplate (tblNovac) | M |
| 17 | Samo Plaćeno/Neplaćeno | Visok | **Dizajnersko ograničenje** — dva stanja (:565-589) | P3 | „Delimično plaćeno" ako biznis zatraži | M |
| 18 | Default PRINT | Visok | **Tačno** — `DocResolveMode(..., "PRINT")` (:495); nameran default | P3 | Prihvatljivo; svestan izbor | — |
| 19 | PDF overwrite / slaba putanja | Visok | **Tačno** — root workbook-a, samo `/` zamena, bez timestampa (:500-501) | P3 | Folder `Fakture\` + timestamp (obrazac prijemnice) | S |
| 20 | Cena 0 bez razloga | Srednji | **Tačno** — dozvoljena (:233-236) uz total>0 (:253-256) | P3 | Opcioni confirm za nultu stavku | S |
| 21 | Klasa/broj prijemnice nisu validirani | Srednji | **Tačno** — sirovo preuzimanje (:225-226) | P3 | Warning za prazne labele | S |
| 22 | Nema valute/poreza/UOM | Srednji | **Dizajnersko ograničenje** — minimalan lokalni model | P3 | Proširenje uz SEF zahteve | L |
| 23 | Datum fakture je uvek danas | Srednji | **Tačno** — `Date` (:263) | P3 | Opcioni validirani datum parametar | S |
| 24 | Item order zavisi od caller kolekcije | Srednji | **Tačno** — `For Each` (:295); UI redosled nameran | P3 | Prihvatljivo | — |
| 25 | Storno status-update je silent no-op | Srednji | **Tačno** — `Exit Sub` (:544-546) | P3 | Vratiti Boolean/status | S |
| 26 | Nema status event history | Srednji | **Tačno** — nema domain eventa u `UpdateFakturaStatus` | P3 | `Monitor_Event` pri promeni statusa | S |
| 27 | `stavka(0)` neformalni Variant ugovor | Srednji | **Tačno** — helper sa jasnom porukom greške (:674-690) | P3 | Prihvatljivo; typed DTO opciono | — |
| 28 | Print uključuje sve matching item redove | Srednji | **Tačno** — bez dedupe/orphan provera (:467-479) | P3 | Guard duplicate PrijemnicaID pri printu | S |
| 29 | Mešani SR/DE error tekstovi | Nizak | **Tačno** — „fehlgeschlagen" (:28, :286, :318) | P3 | Preseliti u modPoruke usput | S |
| 30 | Hardkodovan monitoring user „Operator" | Nizak | **Tačno** — :38, :76 | P3 | `Environ("Username")` | S |

Bilans: 30 — 23 Tačno, 5 Delimično, 2 Dizajnersko; 0×P1, 7×P2, 23×P3.

**Bilans bloka G (133):** 105 Tačno / 18 Delimično / 10 Dizajnersko-Prihvaćeno / 0 Netačno. **3×P1:** cross-tab print target (FM-0030 #1), reprint storniranog otkupa kroz fallback (FM-0031 #3), faktura merge-korupcija (FM-0031 #19). Ključne rekalibracije: revers hidden-state i `DocPrintWs` default nemaju živ pogrešan put; vozački znak je stvarna interna kontradikcija ali bez potrošača; cross-buyer neutralisan UI filtriranjem; status/datum fakture samoizlečiv.

---

## Zbirni bilans (665/665 stavki obrađeno, 0 preskočeno)

| Ocena | Broj | % |
|---|---:|---:|
| **Tačno** (potvrđeno u kodu) | 515 | 77,4% |
| **Delimično** (jezgro tačno, težina/formulacija korigovana) | 78 | 11,7% |
| **Dizajnersko ograničenje / Prihvaćeno** | 63 | 9,5% |
| **Netačno** (opovrgnuto) | 6 | 0,9% |
| **Nije proverivo statički** | 3 | 0,5% |

**Netačne stavke (svih 6):** FM-0006 #6 (avans snapshot pokrivenost — potvrđena kao potpuna); FM-0011 #9 i FM-0012 #15 (`tblPaletaIstorija` ne postoji u kodu); FM-0015 #9 (`Monitor_Event` je interno pod `On Error Resume Next` — ne može podići grešku); FM-0020 #20 (prazan Dictionary `Keys` ne obara petlju); + 1 rekalibracija u FM-0031 #5 klasi (bez živog puta).

**Hitnost (po redovima; nekoliko P1 su isti defekt viđen iz više fajlova):**
- **P0: 1** — pozicioni 17-kolonski insert u `SaveNovac` (modNovac.bas:197-204).
- **P1: 60 redova** (≈52 jedinstvena defekta) — koncentrisani u: izveštajima (11), banci (10), frmDokumenta (5), modAutoHladnjaca (5), storno/invariant lancu (6+2), otkup UI (5), print (3), infra (3).
- **P2: ~200** · **P3/Prihvaćeno: ~400**.

## Najisplativiji paketi popravki (svi verifikovani, pretežno S napor)

1. **Finansijski guard paket:** `RequireColumns` u `SaveNovac` (P0) + target-owner/target-active/no-op guardovi avansa + `StornoNovac`→`UpdateOtkupStatus` (storno isplate danas trajno sakriva dug) + novac storno broj→`NovacID` + stornirane fakture van uplata liste.
2. **„Lažni uspeh" storno lanac (6×S):** modStornoFlow #2 (5 grana bez context guarda), #3 (paletni detach false-success), #7 (ignorisan relink count) + modDokumentInvariant #1 (`0=0` prolazi za nepostojeću zbirnu) i #3 (sum greška postaje validna nula) + modStorno #12 (`LookupActiveID` multi-match).
3. **Hladnjača auto-lanac (5×S):** `outBrPrij` posle uspeha; backfill brojevi po postojećim prijemnicama; provera `otpID`/zbirna rezultata; propagacija link greške.
4. **frmDokumenta unos (5×S):** Kl.II checkbox blokada; smer ambalaže obavezan; malina auto-zbirna vidljiv pad; prefill generacija; live manjak marker.
5. **Izveštaji — poverenje u brojke:** revers štampa (storno filter + tip ambalaže — 2×S); zbirna tab-matrica; status/štampa iz `m_cur*`; isplate kooperanta po stanici; „nema prijema" umesto 0%/100%; kupac per-vrsta raspodela; kartice „Početno stanje".
6. **Banka:** dedupe + broj računa (tihi gubitak transakcije); subscript pad kod 3+ kandidata; smer ručnog mapiranja; Activate auto-map iza dugmeta/vidljiv; preview/command usklađivanje; stale override clamp + finalna saldo revalidacija pre CSV.
7. **Print:** blokada reprint-a storniranog otkupa; `.UnMerge` u `FillFakturaSablon` (korupcija sledeće fakture sa više stavki); `KarticaDetalji_Clear` na promenu taba.
8. **Infra (iz bloka A):** `RollbackTx` garantovan `CleanUp`; `AppendRow` fantomski red; audit snapshot listi po `*_TX` wrapperima; `TBL_CONFIG` van required listi.

## Napomena o FM dokumentu

FM v35 je **činjenično visoko pouzdan**: od 665 stavki samo 6 je opovrgnuto (0,9%), a korekcije su pretežno u kalibraciji težine (multi-user/race klase → P2 u single-writer modelu; „Kritično" za mrtve ili UI-neutralisane puteve → P2/P3). Vredan je nastavka — uz raniju preporuku: ID-jevi nalaza, komit u repo, drift-check, i tok ka KNOWN_ISSUES/ROADMAP.


---
---

# DEO II — Delta trijaža Functional Map v85 (FM-0035…FM-0084)

**Datum:** 2026-07-19
**Izvor:** `AgriX_Functional_Map` v85 — novi unosi FM-0035…FM-0084 (~1007 rizik-stavki/podsekcija).
**Metod:** identičan DEO I (svaka stavka pojedinačno protiv koda, Kritičan/Visok uz citat fajl:linija; 9 paralelnih prolaza). Stari unosi FM-0001…FM-0034 su u v85 **bajt-nepromenjeni** u odnosu na v35 (diff = samo 2 dodate prazne linije) → ostaje validna DEO I trijaža.
**Sidra (dva, jer je FM mešao commit-e):** FM-0035…FM-0075 na `f6313dc` (v2.22.0, verifikovano u worktree kopiji tog commita); FM-0076…FM-0084 na `a0bc9e2` (v2.21.0). Oba proverena protiv tačne kopije koda.

## ⚠️ Napomena o odmaklom `main`-u (utiče na RF-03/RF-04)

FM v85 storno unosi (FM-0011…0015, DEO I, sidro `a0bc9e2`) **ne odražavaju** aktuelni kod: `main`
je od sidra otišao na **v2.24.0** (`58a5075`) kroz storno PR #134–#137, koji su `modStornoFlow.bas`
proširili za **+746 linija** (+ `modDokumentInvariant` +198, `modStorno` +51). Deo P1 nalaza iz
RF-03/RF-04 (context guard u granama, lažni COMPLETED, relink count) je **možda već rešen** tim
PR-ovima. **RF-03/RF-04 se moraju re-verifikovati protiv `origin/main`, ne protiv v35/`a0bc9e2` linija.**
Ostali delta slojevi (SEF, sync, licenca, startup, build) — ti fajlovi se u #134–#137 nisu menjali,
pa delta ostaje validna direktno.

## Zbirni bilans delte (~1007 stavki, 9 blokova)

Dominantno **Tačno** (FM v85 je i dalje činjenično vrlo precizan; opovrgnutih stavki je šačica, uglavnom
sporedni detalji). Glavna korekcija je kalibracija težine: mnogi „Kritično" naslovi padaju na **P2/P3**
zbog (a) single-writer/single-thread modela, (b) reset WithEvents kolekcija koji gasi „stale-click"
scenarije, (c) dokumentovanog fail-open dizajna (auth/licenca/monitoring), (d) mitigacija koje FM
sistematski preskače (strict `GetSingleRowIndexByKey` pre HTTP-a, arhiviran SEF request XML pri retry-ju,
staging-verify-swap, EnableAuth anti-lockout). Veliki deo delte mapira se na **već registrovane** nalaze
(AUD-001, AUD-003, AUD-006, AUD-016, AUD-018, AUD-019, KI-006).

**Novi P0 (1):** SEF klijent mapira HTTP **409 → REJECTED** (`modSEFClient.bas:473-476`) — kod
duplicate/conflict-a faktura se trajno vodi kao odbijena iako dokument postoji na SEF-u; retry šalje isti
`requestId` → korekcioni tok → rizik duple/pogrešne fakture ka poreskoj. Fix S.

**Novi P1 klasteri (~20 jedinstvenih):**
- **SEF correctness:** stornirana faktura je end-to-end poslati­va (validator ne čita `Stornirano`,
  `frmSEF` combo bez filtera, `StornoFaktura` ne dira SEF workflow); qty/price sečeni na 2 decimale →
  aritmetički nekonzistentan UBL (`modSEFMapper`); DueDate < IssueDate uz force-today; fail-soft
  idempotency guard (`HasSuccessfulSEFSubmission` EH→False) dozvoljava dvostruko slanje; stale DocumentID
  kroz resubmit.
- **SEF UX/lifecycle:** `modSEFService` vraća SubmissionID i za REJECTED/TECH_FAILED → `frmSEF` prikazuje
  „Faktura poslata" i za neuspeh; javni `Test_Cancel/Storno…_TX` makroi sa pravnim side-effect-om u Alt+F8;
  blank/unknown status → tiho „SENT"; `frmSEF` combo change ne resetuje prikazani kontekst; recovery vraća
  True i na API failure + lažni „Recovered" event svaki startup za SENDING+remote-terminal.
- **Authorization lanac (van SEF-a, najvažniji):** korisnik sa pravom „Matični podaci" **stiže do Admin
  panela** (Očisti tabele, Migracija, VBA import/export, fleet publish sa šifrom iz `modConfig`) — guard je
  samo na „Korisnici" (`modMaticniLookups.bas:254-259`), a shell propušta `frmStammdaten`
  (`OblastZaFormu`="" → `frmOtkupAPP.frm:1072-1077`); `modAdmin`/`modPodesavanja`/`ShowConfigSheet` bez
  sopstvene provere.
- **Startup/integracija:** `Workbook_Open` **ne poziva `AccessWasDenied`** iako komentar i runbook tvrde da
  poziva (deny se oslanja samo na neproveren `OnTime` close); lažni `STARTUP_SUCCESS` posle deny-ja;
  `frmOtkupAPP.btnBanka_Click` knjiži novac (auto-map na Activate) **pre** auth provere.
- **Self-update:** faza 1 Remove-uje failed `.frm` a faza 2 ga ne uvozi → komponenta nestaje (FM čak
  potcenio); nema manifest completeness check-a (mešana verzija koda).
- **Cenovnik:** stale auto-cena (`If c > 0` ne prazni polje → cena prethodnog proizvoda u `frmOtkup`/`frmDokumenta`).

**Novi P2 paketi (mali delta, veliki efekat):** publish-guard (placeholder/dirty deny + disk↔workbook SHA
cross-check u `PublishReleaseToDrive`); BuildGuard scan plain-range logova (`SETUP_LOG`, test logovi);
environment guard + Boolean `Core` u shipped test suite-ovima (E2E gate danas prijavljuje PASS na svaki
non-throw); empty-source cloud-wipe guard; `SetPWAMasterSyncLock` full-tab overwrite briše `STANICA_LOCK_*`;
atomski rename-par u Sheets swap-u; `modDrive` find-error→duplikat release lanac; SEF `asOfDate`/`Datum` guard
u cenovniku; OAuth OOB migracija + DPAPI za refresh token; PasswordChar za „secret" polja u Podešavanjima.

---

## Delta blok 1 — frmFakturisanje + frmSEF (FM-0035…FM-0036, 52 stavke) [sidro f6313dc]

Audit je kompletan — svih 52 stavke verifikovane protiv koda u worktree-u (`/tmp/claude-0/.../scratchpad/wt-f6313dc/src-vba/`, anchor `f6313dc`). Reference u tabelama: `frmFak` = frmFakturisanje.frm, `frmSEF` = frmSEF.frm, ostalo puni nazivi modula.

### FM-0035 — `frmFakturisanje.frm`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | Confirmation snapshot ≠ commit snapshot | Kritičan | **Tačno** — potvrda iz keša (frmFak:685–696), core ponovo čita tblPrijemnica (modFaktura.bas:154–251); single-writer smanjuje verovatnoću | P2 | frmFak: pre potvrde re-učitati canonical količine/cene iz sveže tabele (mini preview), ili posle create-a prikazati stvarni iznos | M |
| 2 | UI ne prikazuje automatski avans | Kritičan | **Tačno** — `ApplyAvansToFaktura` unutar CreateFaktura (modFaktura.bas:331); potvrda bez avansa (frmFak:693–696) | P2 | frmFak: u confirm poruku dodati raspoloživi avans kupca i napomenu o auto-primeni | S |
| 3 | Auto-selected faktura se štampa bez potvrde | Visok | **Tačno** — auto `ListIndex=0` (frmFak:518–520), direktan `PrintFaktura` (frmFak:774), default mode "PRINT" (modFaktura.bas:495) | P2 | frmFak.btnStampaj: MsgBox potvrda sa brojem/datumom/iznosom pre PrintFaktura | S |
| 4 | Zbirni izbor nije detaljno potvrđen | Visok | **Tačno** — potvrda samo kupac/count/total (frmFak:693–696) | P2 | U potvrdu dodati listu brojeva prijemnica (prvih N + „…još k") | S |
| 5 | Nema max stavki/layout guarda | Visok | **Delimično** — guard ne postoji, ali sablon upisuje sve stavke i štampa multi-page (`FitToPagesTall=False`); realno samo cleanup 81 reda (modPrint.bas:1904) | P3 | modPrint: cleanup range vezati za stvarni prethodni obim; opciono soft-limit upozorenje | S |
| 6 | Modeless forma ne osvežava podatke | Visok | **Delimično** — guard tačan (frmFak:51), ali promena sekcije unload-uje formu (frmOtkupAPP.frm:1105–1108) → povratak = svež load | P3 | Na Activate (kad je m_SetupDone) osvežiti cmbFaktura/uplate | S |
| 7 | Case-sensitive `Da` statusi | Visok | **Tačno** — exact `"Da"` (frmFak:351,388,484,629; modFaktura.bas:662–664; FilterArray `<>` binarno, modArrayUtils.bas:90–91); PrintFaktura ima UCase (modFaktura.bas:435) — nekonzistentno; upisi su app-kanonski | P2 | modHelpers: centralni `JeDa()` (UCase$/Trim$) i zamena na navedenim mestima | S |
| 8 | Stored status i live uplata mogu se razići | Visok | **Delimično** — prikaz stored statusa tačan (frmFak:489,507), ali `UpdateFakturaStatus` se zove pri knjiženju uplata (modBankaMapiranje.bas:291; modDokumenta.bas:1319,1972) → raskorak vanredan | P3 | Pre punjenja cmbFaktura prikazati i „uplaćeno/iznos" ili recalc statusa prikazanih faktura | S |
| 9 | Faktura default izbor prati fizički red | Visok | **Delimično** — obrnut loop (frmFak:483) kod append-only tabele daje najnoviju prvu; ruši se samo ručnim sortiranjem | P3 | SortArray po COL_FAK_DATUM desc pre punjenja | S |
| 10 | Print nema result/confirmation | Visok | **Tačno** — `PrintFaktura` je Sub, tihi izlaz kad šablon vrati Nothing (modFaktura.bas:492), forma bez ikakvog feedback-a (frmFak:774) | P2 | Nothing → Err.Raise; posle štampe status poruka (mode/putanja PDF) | S |
| 11 | Payment detalj zaokružuje na 0 decimala | Srednji | **Tačno** — `"#,##0"` (frmFak:409–411) | P3 | Format "#,##0.00" | S |
| 12 | Status suma locale string parsing | Srednji | **Tačno** — Replace `.`/`,` + CDbl nad formatiranim tekstom (frmFak:170–175) | P3 | Sumirati numeric vrednosti iz m_PrijemniceData preko m_DataIndices | S |
| 13 | UI ne prikazuje buyer ownership check | Srednji | **Tačno** — CreateFaktura ne proverava KupacID prijemnice (modFaktura.bas:184–241); jedina zaštita UI filter (AUD-011) | P2 | modFaktura.CreateFaktura: validirati COL_PRJ_KUPAC = kupacID po stavci (AUD-011 fix u core) | S |
| 14 | `m_IsLoading` neaktivan guard | Srednji | **Tačno** — deklarisan (frmFak:39), čitan (frmFak:222), nigde se ne postavlja na True (grep prazan) | P3 | Postaviti oko FillComboDisplayID u Activate, ili obrisati | S |
| 15 | Setup flag preuranjen | Srednji | **Tačno** — `m_SetupDone = True` pre setup-a (frmFak:51–52) | P3 | Flag postaviti na kraj uspešnog setup-a | S |
| 16 | Nema business sort-a stavki | Srednji | **Tačno** — GetPrijemniceByKupac/FilterArray zadržava fizički red (modDokumenta.bas:1205–1250) | P3 | SortArray po datumu/broju pre punjenja liste | S |
| 17 | Variant stavka sadrži ignorisane podatke | Srednji | **Tačno** — core koristi samo `stavka(0)` (modFaktura.bas:678, komentar 172–174); namerni defense-in-depth | P3 | Slati samo kolekciju PrijemnicaID-jeva (promena potpisa) ili dokumentovati ugovor | M |
| 18 | Create success prikazuje tehnički ID | Srednji | **Tačno** — `"Faktura kreirana: " & fakturaID` (frmFak:706) | P3 | LookupValue COL_FAK_BROJ i prikaz poslovnog broja | S |
| 19 | Fill errors izgledaju kao no-data | Srednji | **Tačno** — EH samo LogErr + Clear (frmFak:524–527); štampa javlja „Nema faktura" (frmFak:762) | P3 | U EH prikazati poruku o grešci (kao btnUnesi EH) | S |
| 20 | SEF nije prefiltriran selected fakturom | Srednji | **Tačno** — btnSEF otvara generički (frmFak:786–796); frmSEF nema init API | P3 | frmSEF: public init sa FakturaID + poziv iz btnSEF | S |
| 21 | Hardkodovan RSD | Srednji | **Dizajnersko ograničenje** — ceo faktura model je single-currency domaći | Prihvaćeno | Ništa sada; eventualno konstanta u modConfig | S |
| 22 | Refresh pozivi duplirani | Nizak | **Tačno** — btnUnesi_Click već zove FillFaktureZaKupca (frmFak:299), pa opet (frmFak:709) | P3 | Ukloniti drugi poziv | S |
| 23 | Mešani SR/DE komentari | Nizak | **Tačno** — „Rechnungserstellung", „PRIJEMNICE LADEN", „DRUCKEN" | P3 | Ništa hitno; postepeno uskladiti pri izmenama | S |

**Bilans FM-0035:** 18 Tačno / 4 Delimično / 0 Netačno / 1 Dizajnersko ograničenje; hitnost: 0×P0, 0×P1, 7×P2 (#1,2,3,4,7,10,13), 15×P3, 1×Prihvaćeno.

### FM-0036 — `frmSEF.frm`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | Combo selection ≠ prikazani context | Kritičan | **Tačno** — nema `cmbFaktura_Change`; akcije čitaju live combo (frmSEF:279–283), labele/dugmad samo na btnUcitaj; wheel-hook povećava rizik slučajne promene | P1 | Dodati cmbFaktura_Change → ClearSEFInfo (reset labela i dugmadi) | S |
| 2 | Prazan SubmissionID daje success | Kritičan | **Delimično** — prazan ID bez exceptiona praktično nedostižan (raise, modSEFService.bas:42–45); ALI REJECTED/TECH_FAILED putanje vraćaju submissionID (modSEFService.bas:242–280,384) → „Faktura poslata" (frmSEF:458) i za neuspelo slanje | P1 | Posle send-a pročitati SEFWorkflowState i granati poruku (poslata/odbijena/tehnička greška) | S |
| 3 | `DoEvents` reentrancy | Kritičan | **Delimično** — prozor samo u jednom DoEvents pre čitanja ID-ja (frmSEF:439–442); tokom sinhronog HTTP-a nema pump-a, MsgBox modalno blokira; queued klikovi drugih dugmadi ipak mogu ući | P2 | Globalni mBusy + disable svih dugmadi/comba u svim akcijama; ukloniti DoEvents | S |
| 4 | Partial-info stale mix | Kritičan | **Delimično** — LookupValue nikad ne raise-uje (modDataAccess.bas:439–467) → labele se uvek sve postave; jedini mid-raise je event-schema subscript (frmSEF:360–369) → star ostaje samo button-state, uz vidljiv error MsgBox | P3 | ClearSEFInfo na početku load-a + RequireColumnIndex u event load-u | S |
| 5 | Sve fakture bez lifecycle filtera | Visok | **Tačno** — svi redovi bez filtera (frmSEF:264–275) | P2 | Preskočiti Stornirano="Da" (UCase) i terminalne STORNO/CANCELLED | S |
| 6 | Nema exact-single FakturaID guarda | Visok | **Tačno** — combo svaki red, LookupValue prvi pogodak; `RequireSingleFakturaRow` postoji (modFaktura.bas:609) ali se ovde ne koristi; dupli ID = ručno oštećenje | P3 | U LoadSelectedFakturaInfo FindRows + upozorenje za count>1 | S |
| 7 | Button state dve state dimenzije | Visok | **Delimično** — razdvojenost tačna (frmSEF:395–418), ali uslovi ogledaju servisne validatore (cancel DRAFT/NEW/ERROR; storno SENT/ACCEPTED/REJECTED = modSEFValidator.bas:333–370) — dizajn prati SEF model | P3 | Opciono invariant-warning label za nekombinabilne workflow/status parove | S |
| 8 | Lokalni storno se ne proverava | Visok | **Tačno, i šire** — ni UI ni `ValidateFakturaForSEF` (modSEFValidator.bas:58–165, nema Stornirano); lokalni storno fakture postoji (modStorno.bas:702–713) i NE dira SEFWorkflowState → stornirana LOCAL_FINALIZED faktura je poslavljiva na SEF | P1 | Stornirano check u ValidateFakturaForSEF (core) + filter u combu (#5) | S |
| 9 | Send confirmation samo tehnički ID | Visok | **Tačno** — „Poslati fakturu FAK-x…" (frmSEF:449) | P2 | U confirm dodati broj/kupca/iznos (LookupValue) | S |
| 10 | Refresh/resubmit/recovery result-less | Visok | **Delimično** — sve raise-uju na neuspeh (EH preskače success poruku); resubmit validira REJECTED (modSEFValidator.bas:396–400); refresh vraća Boolean koji forma ODBACUJE (modSEFStatusSync.bas:27 vs frmSEF:482); labele se osveže pre poruke | P3 | Prikazati novi status u poruci; iskoristiti Boolean rezultat refresh-a | S |
| 11 | Batch bez confirmation/preview/result | Visok | **Tačno** za UI — bez potvrde/rezultata (frmSEF:612–638); servisi broje Found/Recovered/Failed samo u Monitor, per-item greške progutane → poruka „uspeh" i kad deo padne (modSEFStatusSync.bas:461+; modSEFService.bas:714+) | P2 | Servisi da vrate summary (processed/changed/failed); forma confirm pre + prikaz posle | M |
| 12 | Single recovery bez confirmation | Visok | **Delimično** — potvrde nema (frmSEF:590–604, jedina akcija bez), ali guarded (raise ako nije SENDING) i bezbedna po dizajnu (DocumentId→refresh; inače TECH_FAILED za retry iste submisije, modSEFService.bas:641–689) | P3 | Dodati MsgBox potvrdu radi konzistentnosti | S |
| 13 | Version se ne koristi za CAS | Visok | **Delimično** — tačno da se samo prikazuje (frmSEF:320), ali single-writer desktop nema konkurentne instance | P3 | Ništa sada; uz multi-user uvesti verziju u action guard | L |
| 14 | SEFDocumentID nije deo action guarda | Visok | **Tačno** za UI (frmSEF:414–416), ali servis raise-uje jasnu grešku bez DocumentId (modSEFValidator.bas:322–330,354–363) → vidljiva poruka, ne tiha greška | P3 | U enable uslov dodati Len(SEFDocumentId)>0 | S |
| 15 | Forma se ne osvežava na Activate | Visok | **Delimično** — guard tačan (frmSEF:31–32), ali promena sekcije unload-uje formu (frmOtkupAPP.frm:1105–1108) → povratak = svež load; stale samo unutar iste sesije forme | P3 | Na Activate posle setup-a reload combo liste | S |
| 16 | Combo zavisi od `.frx` kolona | Srednji | **Tačno** — bez ColumnCount/Widths/Bound u kodu (frmSEF:262–275); frmFakturisanje ih postavlja (frmFak:126–132) | P3 | Postaviti svojstva u kodu kao u frmFakturisanje | S |
| 17 | Nema preselection iz fakturisanja | Srednji | **Tačno** — isto kao FM-0035 #20 | P3 | Public init metoda + poziv iz frmFakturisanje | S |
| 18 | Nema sortiranja/filtera/pretrage | Srednji | **Tačno** — fizički red, bez pretrage (frmSEF:272–275) | P3 | Bar obrnuti red (najnovije prvo) kao u frmFakturisanje | S |
| 19 | Event schema nije fail-fast | Srednji | **Tačno** — GetColumnIndex bez provere 0 → `data(i,0)` subscript (frmSEF:360–369) | P3 | RequireColumnIndex za 4 event kolone | S |
| 20 | Event log bez full-text/copy/export | Srednji | **Tačno** — samo ListBox kolone | P3 | DblClick na red → MsgBox/textbox sa punim Details | S |
| 21 | Status `ERROR` nema error boju | Srednji | **Tačno** — Case Else → TXT_LIGHT (frmSEF:327–338), a cancel enabled baš za ERROR (frmSEF:414) | P3 | `Case "ERROR"` → CLR_ERROR() | S |
| 22 | InputBox reason nestrukturiran | Srednji | **Tačno** — InputBox, samo nonempty (frmSEF:530–531,566–567) | P3 | Min dužina + Trim; reason code opciono | S |
| 23 | Storno broj nije validiran | Srednji | **Tačno** u UI — slobodan opcion tekst (frmSEF:569); pravila na servisu | P3 | Osnovna format provera pre poziva servisa | S |
| 24 | Last error može biti odsečen | Srednji | **Nije proverivo statički** — Label svojstva (WordWrap/AutoSize) žive u binarnom `.frx` | P3 | Klik na lblLastError → MsgBox pun tekst | S |
| 25 | Combo schema failure kao no-data | Srednji | **Tačno** — tihi `Exit Sub` bez loga/poruke (frmSEF:265,270) | P3 | LogErr + poruka „šema tblFakture neispravna" | S |
| 26 | Help ne pokriva sve workflow-e | Srednji | **Tačno** — samo READY/SENDING/SENT/ACCEPTED/REJECTED (frmSEF:682–695) | P3 | Dopuniti help (TECH_FAILED, SYNC_ERROR, cancel/storno, recovery) kroz modPoruke | S |
| 27 | Nema role/environment indikatora | Srednji | **Delimično** — ulaz u sekciju JE auth-gated opt-in (frmOtkupAPP.frm:1072–1077, modAuth); nema per-akcija prava ni SEF environment prikaza | P3 | Label sa SEF API URL/environment iz konfiga u headeru | S |
| 28 | Tehnički status stringovi direktno | Nizak | **Tačno** — lblWorkflow prikazuje sirove konstante (frmSEF:317) | P3 | Mapiranje na čitljive nazive kroz Poruka katalog | S |
| 29 | Mešani SR/EN captioni | Nizak | **Tačno** — „Recover sending", „Retry slanje na SEF" | P3 | Ništa hitno; uskladiti kroz modPoruke | S |

**Bilans FM-0036:** 19 Tačno / 9 Delimično / 0 Netačno / 1 Nije proverivo statički; hitnost: 0×P0, 3×P1 (#1,2,8), 4×P2 (#3,5,9,11), 22×P3.

**Ključne korekcije/nadgradnje FM-a otkrivene auditom:**
1. **FM-0036 #2:** stvarni false-success mehanizam nije prazan SubmissionID (nedostižan — servis raise-uje), nego to što REJECTED/TECH_FAILED ishodi komituju i vraćaju ID → modal „Faktura poslata" i za neuspešno slanje (modSEFService.bas:242–280,384 + frmSEF:454–458).
2. **FM-0036 #8 je ozbiljniji nego što FM tvrdi:** lokalni storno fakture postoji (modStorno.bas:702–713), ne dira SEF workflow, a `ValidateFakturaForSEF` ne proverava Stornirano — servis dakle NE odbija slanje lokalno stornirane fakture; to je jedini nalaz gde poslednja linija odbrane ne postoji.
3. **FM-0036 #4/#15 i FM-0035 #6 preuveličani:** `LookupValue` nikad ne raise-uje (vraća Empty), a navigacija između sekcija unload-uje sadržajnu formu (frmOtkupAPP.frm:1105–1108) pa se na povratku podaci sveže učitavaju.

Ukupno 52/52 stavke: **37 Tačno / 13 Delimično / 0 Netačno / 1 Dizajnersko / 1 Nije proverivo**; hitnost: 3×P1, 11×P2, 37×P3, 1×Prihvaćeno. Nijedan nalaz nije P0 (nema korupcije podataka — svi upisi idu kroz TX servise koji raise-uju/rollback-uju), a sva tri P1 su u frmSEF/SEF servisnom sloju i rešiva malim deltama (S).

---

## Delta blok 2 — SEF service, status sync, persistance (FM-0037…FM-0039, 78 stavki) [sidro f6313dc]

# Audit rizik-nalaza FM-0037…FM-0039 protiv koda (worktree `wt-f6313dc/src-vba/`)

Skraćenice u citatima: **Svc**=`modSEFService.bas`, **Sync**=`modSEFStatusSync.bas`, **Pers**=`modSEFPersistance.bas`, **Cli**=`modSEFClient.bas`, **Val**=`modSEFValidator.bas`, **frm**=`frmSEF.frm`, **Main**=`modMain.bas`, **Tx**=`clsTransaction.cls`, **Cfg**=`modConfig.bas`. Ključni kontekst potvrđen u kodu: klijent **nikad ne vraća `Nothing`** (EH grane konstruišu failure objekat: Cli:52-65, 111-125, 189-200, 265-276); retry šalje **isti `requestId`** SEF-u kao query parametar (Cli:411); startup recovery se poziva iz Main:100; `RollbackTx` je idempotentan (Tx:84); kalibracija single-writer desktop.

### FM-0037 — `modSEFService.bas` (28 redova)

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | Send return nije success (Rejected/TECH_FAILED vraćaju SubmissionID) | Kritičan | **Tačno** — Svc:384 vraća `submissionID` u sve 4 grane; frm:454-458 bezuslovno prikazuje „Faktura poslata“ i za REJECTED/TECH_FAILED | P1 | frm posle poziva čita `GetFakturaSEFWorkflowState` (već radi `LoadSelectedFakturaInfo`) i poruku bira po ishodu; dugoročno strukturiran rezultat | S |
| 2 | Recovery bez DocumentID-a ne query-je SEF → moguć retry već prihvaćene fakture | Kritičan | **Delimično** — grana potvrđena (Svc:652-686 → TECH_FAILED → reuse Svc:35-54), ali retry šalje **isti requestId** (Svc:37,167 + Cli:411 `?requestId=`) — SEF dedup ključ postoji; stroga garancija servera **nije proveriva statički** | P2 | Na demo SEF-u testirati ponovni POST istog requestId; opciono remote lookup pre retry-ja | M |
| 3 | Nema stuck-age/lease — aktivno slanje može biti recovery-jano | Kritičan | **Delimično** — kriterijum zaista ne postoji (Svc:757), ali single-writer + VBA single-thread: na startup-u (Main:100) svaki SENDING JE stvarno prekinut; paralelna instanca van kalibracije | Prihvaćeno | Opciono `SendingStartedAt` kolona radi vidljivosti | S |
| 4 | Nema CAS/cross-user claim — dve instance šalju istu fakturu | Kritičan | **Dizajnersko ograničenje** — tačno da nema CAS (Svc:27-28), ali single-writer; usput `UpdateFakturaSEFState_Row` u TX-u ponovo čita stanje i validira tranziciju (Pers:117-121), što blokira ilegalne preskoke | Prihvaćeno | Dokumentovati single-writer pretpostavku | S |
| 5 | Cancel/storno HTTP pre durable intent-a | Kritičan | **Tačno** — HTTP na Svc:461/540 PRE `BeginTx` na Svc:463/542; pad posle remote uspeha ne ostavlja trag | P2 (ne P0: oporavak postoji — DocumentID je poznat, `RefreshSEFStatus_TX` povlači CANCELLED/STORNO) | Upisati intent event pre HTTP-a ili posle greške naložiti/pokrenuti refresh | S |
| 6 | Public produkcioni test makroi (send/cancel/storno) | Kritičan | **Tačno** — Svc:899-1122; `Test_CancelInvoiceOnSEF_TX` (Svc:940-949, „FAK-00007“) i `Test_StornoInvoiceOnSEF_TX` (Svc:951-960) bez potvrde, vidljivi u Alt+F8; pravni side effect | P1 | Premestiti u postojeći `modSEFTests.bas` ili guard (`SEF_ENV`=DEMO + potvrda) | S |
| 7 | SEFStatus sadrži lokalna READY/SENDING stanja | Visok | **Tačno** — Svc:84 (`sefStatus:=WF_SEF_READY`), Svc:131 (SENDING), Svc:674 (TECH_FAILED u recovery-ju); krši komentar Svc:5-8/Sync:20 | P3 (validatori cancel/storno rade nad SEFStatus, ali zagađene vrednosti nisu u allowed listama → fail-closed) | U PREP/TX1 ne prosleđivati `sefStatus` (ostaviti postojeći) | S |
| 8 | `HTTP_SENT` event pre HTTP-a | Visok | **Tačno** — Svc:136-141 u TX1, HTTP tek Svc:167; konstanta `"HTTP_SENT"` Cfg:683; tekst poruke („submission started“) je ipak tačan | P3 | Novi event tip `SUBMISSION_PREPARED` | S |
| 9 | Payload nije vezan za invoice CAS/version | Visok | **Dizajnersko ograničenje** — build (Svc:58) i HTTP (Svc:167) su u istom sinhronom pozivu; single-writer ne dozvoljava izmenu fakture između | Prihvaćeno | — | — |
| 10 | PREP TX ostavlja READY posle kasnijeg failure-a | Visok | **Delimično** — stanje potvrđeno (Svc:74-107 zaseban commit), ali READY je regularno sendable stanje (Val:132, Val:14-15) — ponovni klik nastavlja tok | P3 | Event poruku dopuniti „slanje nije započeto“ | S |
| 11 | Cancel/storno ne menjaju local workflow | Visok | **Tačno** — Svc:468-481/547-560 samo `UpdateFakturaSEFRefreshFields_Row`; `WF_SEF_STORNO` postoji (Cfg:660) i tranzicije SENT/ACCEPTED→STORNO su dozvoljene (Val:26,38), ali ga **niko ne postavlja** (grep: samo validator/testovi). Dupli storno ipak blokiran preko SEFStatus (Val:366-373) | P2 | Na storno uspeh postaviti workflow `SEF_STORNO` (tranzicija već dozvoljena); za cancel definisati stanje | S |
| 12 | Cancel/storno response `Nothing` nije guardovan | Visok | **Delimično** — guard zaista ne postoji (Svc:468/547), ali klijent po konstrukciji uvek vraća objekat (EH: Cli:189-200/265-276) — premisa trenutno nedostižna | P3 | Defanzivni `If response Is Nothing Then Err.Raise` | S |
| 13 | Recovery refresh rezultat nije proveravan → lažni recovery event | Visok | **Tačno** — Svc:653-662 bezuslovni „Recovered“; dokazan konkretan slučaj: remote STORNO/CANCELLED + workflow SENDING → Sync:128-135 menja samo refresh polja → faktura **ostaje SENDING**, event tvrdi recovery, ponavlja se svaki startup; a API-failure grana (Sync:165-170) pokušava SENDING→SYNC_ERROR što Val:17-22 zabranjuje → izuzetak, faktura ostaje SENDING | P2 | Posle refresh-a proveriti `workflow <> SENDING`; definisati mapiranje SENDING+terminal remote status | M |
| 14 | Recovery event van refresh TX-a | Visok | **Tačno** — refresh ima svoj TX (Sync:51-183), append Svc:655-660 posle njega; pad append-a = commitovan refresh + prijavljen failure | P3 | Append pod tolerantnim error handling-om ili rezultat-aware | S |
| 15 | Batch result ne postoji | Visok | **Tačno** — `Sub` (Svc:714), brojači lokalni (Svc:720-722); frm:629-632 fiksna poruka | P3 | Vratiti summary (found/recovered/failed) i prikazati | S |
| 16 | Batch svako SENDING smatra stuck | Visok | **Delimično** — isto kao #3 (Svc:757); na startup-u tačno po dizajnu | Prihvaćeno | — | — |
| 17 | Duplicate FakturaID se obrađuje više puta | Visok | **Delimično** — fizička petlja bez dedupe (Svc:752-777), ali duplikat PK = već korumpirani podaci; drugi prolaz pada glasno (Svc:643-646), a write strana raise-uje duplikate (Pers:427-430) | P3 | Dedupe kolekcija u petlji | S |
| 18 | VersionNo nije optimistic token | Visok | **Dizajnersko ograničenje** — single-writer; paralelno računanje verzije nedostižno | Prihvaćeno | — | — |
| 19 | AttemptCount je stalno 0 | Visok | **Tačno** — svi `Monitor_SEF` pozivi `attemptCount:=0` (Svc:157,306,323,340,357,374,420,770,800,829) | P3 | Izvesti broj iz count-a submission redova fakture | S |
| 20 | BusinessInvoiceNo je FakturaID | Visok | **Tačno** — `businessInvoiceNo:=fakturaID` svuda (Svc:153,301,317…) | P3 | Jednom pročitati `BrojFakture` i prosleđivati | S |
| 21 | `apiStatus`/response flagovi nisu validirani | Srednji | **Delimično** — servis veruje flagovima (Svc:197,242), ali klijent u svim granama postavlja konzistentan skup i neprazan `apiStatus` (Cli:450-505, EH grane) | P3 | Invariant assert (tačno jedna finalna kategorija) | S |
| 22 | LastSyncAt se menja i pri neuspešnom cancel/storno | Srednji | **Tačno** — Svc:496/575 van `If response.Success` grane | P3 | Preimenovati semantiku ili dodati `LastSuccessfulSyncAt` | S |
| 23 | Batch partial failure se zove SUCCESS | Srednji | **Delimično** — finalni event uvek `..._SUCCESS`/INFO (Svc:843-848), ali per-item `SEF_RECOVERY_INVOICE_FAIL`/CRITICAL postoji (Svc:792-804) — alerting ima signal | P3 | Severity WARN kad `failedCount>0` | S |
| 24 | Hardkodovan user Operator | Srednji | **Delimično** — samo u `Monitor_Event` (Svc:729,849,886); event log već koristi stvarnog korisnika (`GetCurrentOperatorName`, Pers:489-501) | P3 | Proslediti isti helper u monitoring | S |
| 25 | Request body retention/security nije rešena | Srednji | **Dizajnersko ograničenje** — XML u `tblSEFSubmission` (Svc:124); workbook ionako sadrži sve podatke faktura, XML ne dodaje novu klasu tajni | Prihvaćeno | Po potrebi kasnija arhiva/čišćenje starih body-ja | M |
| 26 | Event taxonomy je neprecizna | Srednji | **Tačno** — `HTTP_SENT` pre HTTP-a (Svc:139), generički `SYNC_OK` za cancel (Svc:479) i storno (Svc:558) | P3 | Posebni event tipovi (CANCELLED_ON_SEF, STORNO_ON_SEF…) | S |
| 27 | EH duplira rollback | Srednji | **Tačno** — doslovno dupliran blok Svc:427-436 + dupli `LogErr` (Svc:398,432); bezopasno jer je `RollbackTx` idempotentan (Tx:84 `If Not mActive Then Exit Sub`) | P3 | Obrisati drugi blok | S |
| 28 | Testovi su smoke bez assertions | Srednji | **Delimično** — za ovaj modul tačno (samo `Debug.Print`), ali `modSEFTests.bas` ima assertion infrastrukturu (`AssertTrue`, `LogPass/LogFail`) | P3 | Konsolidovati smoke procedure u `modSEFTests` | S |

**Bilans:** 28/28 provereno; 14 Tačno, 10 Delimično, 4 Dizajnersko ograničenje, 0 Netačno. Hitnost: **P1×2** (#1 lažna „Faktura poslata“, #6 javni cancel/storno makroi), **P2×4** (#2, #5, #11, #13), P3×16, Prihvaćeno×6. Najvredniji novi dokaz: #13 — SENDING + remote STORNO/CANCELLED ostaje trajno SENDING uz lažni „Recovered“ event svaki startup.

### FM-0038 — `modSEFStatusSync.bas` (27 redova)

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | `SYNC_ERROR → SENT` za remote STORNO/CANCELLED | Kritičan | **Delimično** — kod potvrđen (Sync:120-127), ali kombinacija `SENT`+`STORNO` je **eksplicitno dokumentovana kao legalna** (komentar Sync:15); SEFStatus=STORNO se upiše → batch je preskače (Sync:513), a send/storno akcije su blokirane validatorima; SYNC_ERROR je ionako ne-terminalan | P3 | Dvokorak `SYNC_ERROR→SENT→SEF_STORNO` (SENT→STORNO već dozvoljen, Val:26) po ugledu na ACCEPTED/REJECTED hop (Sync:413-424) | S |
| 2 | Unknown successful status → `SEF_SENT` (fail-open) | Kritičan | **Tačno** — Sync:144-159 `Case Else` → target SENT, event „non-final“; klijent prosleđuje raw nepoznat status (Cli:598-600); novi terminalni SEF status ostaje večno „pending“ | P2 | `Case Else`: sačuvati raw status, NE menjati workflow, event `UNKNOWN_STATUS` + manual review; na demo SEF-u proveriti stvarni rečnik statusa | S |
| 3 | Refresh vraća True i na API failure | Kritičan | **Tačno** — Sync:311 bezuslovno posle commita (i za SYNC_ERROR granu Sync:163-179); frm:482-485 prikazuje „status osvežen“ ne gledajući rezultat; Svc recovery isto ignoriše | P2 | Else grana → `False` (ili enum); frm poruka po rezultatu; batch broji po rezultatu | S |
| 4 | SEFStatus čuva lokalni `SYNC_ERROR` | Kritičan | **Tačno** — Sync:168 upisuje `WF_SEF_SYNC_ERROR` u SEFStatus; krši sopstveni model (Sync:20); gubi se poslednji poznati remote status (samoizlečivo sledećim uspešnim poll-om) | P3 | U failure grani ne dirati SEFStatus (samo error polja + workflow) | S |
| 5 | Nema CAS/version zaštite | Kritičan | **Dizajnersko ograničenje** — single-writer kalibracija; stale-response overwrite zahteva paralelnu instancu | Prihvaćeno | — | — |
| 6 | Nema exact-single FakturaID guarda | Visok | **Delimično** — read je first-match (`LookupValue` kroz Pers:752-787), ali write raise-uje duplikate PRE upisa (Pers:427-430) → fail-late, ne pogrešan upis | P3 | Zajednički exact-single read helper | S |
| 7 | Response DocumentID se ne reconciliuje | Visok | **Delimično** — poređenje ne postoji (Sync:69,86,107), ali klijent echo-uje traženi ID (Cli:99) i parser ga menja samo ako body vrati `InvoiceId` (Cli:~545) — mismatch zahteva anomalan server response | P3 | `If response.sefDocumentId <> requested` → warning event | S |
| 8 | Prazan successful `apiStatus` je regularan pending | Visok | **Tačno** — maska potvrđena, i to već u klijentu: Cli:600 `FirstNonEmpty(statusValue, "SENT")` pretvara prazan status u „SENT“ — nerazlučivo od pravog SENT | P3 | Prazan/missing `Status` ključ tretirati kao parser grešku (TL-001 srodno) | S |
| 9 | Response flagovi nisu validirani | Visok | **Delimično** — prioritet grana potvrđen (Sync:63-97), ali `ParseStatusResponse` u svakoj grani postavlja međusobno konzistentne flagove (Cli:503-602) | P3 | Invariant assert | S |
| 10 | Local STORNO/eligibility nije proverena | Visok | **Delimično** — tačno da nema provere, ali refresh je remote read-only + reconciliation upis; osvežavanje lokalno stornirane fakture je često poželjno | P3 | Po potrebi skip lista uz DQ event | S |
| 11 | Batch „Refreshed“ uključuje remote failure | Visok | **Tačno** — Sync:561 broji svaki no-exception poziv, a funkcija vraća normalno i na API failure (Sync:311); rezultat se i ne čita (Sync:519) | P3 | Brojati po povratnoj vrednosti (posle #3) | S |
| 12 | Duplicate FakturaID se osvežava više puta | Visok | **Delimično** — fizička petlja (Sync:502-571); preduslov korumpiran PK; upis pada glasno na duplikatu | P3 | Dedupe | S |
| 13 | Paralelni batch nema lease/claim | Visok | **Dizajnersko ograničenje** — single-writer | Prihvaćeno | — | — |
| 14 | Nema polling backoff/NextPollAt | Visok | **Delimično** — tačno (nema `NextPollAt`), ali batch je isključivo ručna akcija (frm:615, bez schedulera) sa fiksnih 2 s (Sync:566) | P3 | `NextPollAt` tek ako se uvede automatski polling | M |
| 15 | Batch nema result contract | Visok | **Tačno** — `Sub` (Sync:461), brojači lokalni; frm:615-618 fiksna poruka | P3 | Summary povratna vrednost + prikaz | S |
| 16 | Finalni batch INFO pri failure-u | Visok | **Delimično** — summary uvek INFO (Sync:574-586), ali per-item `..._INVOICE_FAIL`/ERROR postoji (Sync:544-557) | P3 | Severity po `failedCount` | S |
| 17 | `LastSyncAt` znači attempt, ne success | Srednji | **Tačno** — Sync:181 u svim granama (i SYNC_ERROR) | P3 | Razdvojiti attempt/success timestampe | S |
| 18 | Pending event pri svakom no-change refresh-u | Srednji | **Tačno** — Sync:111-116 `SYNC_OK` „unchanged (pending)“ na svaki poll | P3 | Event samo na promenu statusa | S |
| 19 | Submission snapshot je nepotreban | Srednji | **Tačno** — Sync:54 snapshotuje `tblSEFSubmission`, a save je zakomentarisan (Sync:57-61) | P3 | Ukloniti snapshot ili vratiti save (odlučiti model, obrisati mrtvi komentar) | S |
| 20 | Empty workflow se tiho inicijalizuje | Srednji | **Tačno** — Sync:372-381 direktan upis target stanja bez DQ eventa | P3 | Dodati data-quality event | S |
| 21 | Final-to-final konflikt nema lokalno pravilo | Srednji | **Delimično** — lokalno pravilo zaista ne postoji, ali `ValidateAllowedTransition` hard-blokira npr. ACCEPTED→REJECTED (Val:37-38) → glasna greška, ne tihi overwrite | P3 | Konflikt logovati kao poseban KONFLIKT/manual-review event umesto generičke greške | S |
| 22 | Fixed Wait blokira UI | Srednji | **Tačno** — `Application.Wait` +2 s po fakturi (Sync:566), bez progress/cancel | P3 | DoEvents/progress, opcioni cancel | M |
| 23 | Prazna tabela nema END monitoring event | Srednji | **Tačno** — Sync:483 `Exit Sub` pre summary-ja (START već poslat Sync:468) | P3 | Skok na summary umesto `Exit Sub` | S |
| 24 | Prazan FakturaID ide do API helpera | Srednji | **Delimično** — tačno da nema pre-provere, ali refresh odmah pada sa jasnom porukom „No SEFDocumentId“ (Sync:40-43) kao item failure | P3 | Preskočiti prazne ID redove uz DQ event | S |
| 25 | Hardkodovan Operator/correlation ID | Srednji | **Tačno** — Sync:472,477,581,586 | P3 | Stvarni korisnik + unique run ID | S |
| 26 | Public test makroi menjaju realno stanje | Srednji | **Tačno** — Sync:630-673 (FAK-00008); remote poziv je read-only, lokalni upis je reconciliation — znatno benignije od send/cancel testova | P3 | Premestiti u `modSEFTests` | S |
| 27 | Dupliran/leading whitespace | Nizak | **Tačno** — Sync:2 ` Option Explicit` (i Svc:627 leading space) | P3 | Kozmetika uz sledeću izmenu | S |

**Bilans:** 27/27 provereno; 15 Tačno, 10 Delimično, 2 Dizajnersko ograničenje, 0 Netačno. Hitnost: P1×0, **P2×2** (#2 fail-open nepoznat status, #3 True-na-failure + frm poruka), P3×23, Prihvaćeno×2. Dve od pet „Kritičan“ ocena FM-a su precenjene: #1 je dokumentovano-legalna kombinacija sa blokiranim akcijama, #4 je samoizlečiv display gubitak.

### FM-0039 — `modSEFPersistance.bas` (23 reda)

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | REJECTED ostaje `PoslatNaSEF=Ne` | Kritičan | **Tačno** — Pers:149-159: „Da“+`SEFSentAt` samo za SENT/ACCEPTED, SENDING vraća „Ne“, REJECTED izostavljen; kolona trenutno **nema nijednog čitaoca u repo-u** (grep VBA+gas+src = 0), ali je vidljiva operateru u tabeli | P2 | I za `WF_SEF_REJECTED` upisati „Da“+`SEFSentAt` (zahtev je stigao do SEF-a) | S |
| 2 | Split-brain last submission (pointer vs CreatedAt sort) | Kritičan | **Delimično** — dva izvora potvrđena (pointer Pers:25-28 vs sort Pers:609-617; oba u `ShouldReuseLastSubmission`, Svc:611-614), ali create+pointer su u istoj TX (Svc:118-134) pa divergencija realno samo kod CreatedAt tie (sekundna rezolucija, single-writer) | P3 | `GetLastSEFSubmissionStatus` čitati po pointer ID-ju umesto sortiranja (rešava i #FM red o tie-u) | S |
| 3 | Read nema exact-single invariant | Kritičan | **Tačno** kao asimetrija — read `LookupValue` first-match (Pers:763), write `GetSingleRowIndexByKey` raise-uje duplikate (Pers:408-433); posledica je **fail-late** (write pukne pre HTTP-a u PREP TX), ne pogrešan upis | P3 | Exact-single i za read helpere | S |
| 4 | Invalid version se tiho resetuje | Kritičan | **Delimično** — prazno→1 je legitiman prvi put; samo neparsabilna vrednost je tihi reset (Pers:40-46, 65-71); verzija je informativna, bez unique constraint-a | P3 | Raise za neparsabilnu (ne-praznu) vrednost | S |
| 5 | Public multi-cell helper bez TX-a | Kritičan | **Tačno** — Pers:88-171 bez `BeginTx`; svi produkcioni pozivi jesu u TX, ali naked-call obrazac postoji (komentar Svc:3 pokazuje Immediate poziv) | P3 | Komentar-ugovor „samo unutar TX“ + interna konvencija; opcioni tx-depth guard | S |
| 6 | Fail-soft idempotency read | Kritičan | **Tačno** — lanac potvrđen: `GetSEFSubmissionsForFaktura` EH→`Empty` (Pers:528-531) → `HasSuccessfulSEFSubmission`=False (Pers:570-572) → duplicate guard prolazi (Val:156-159); preduslov je read/schema kvar | P2 | EH u list-read helperima da raise-uje (fail-closed) bar za `HasSuccessful...` put | S |
| 7 | SENDING briše ever-sent indikator | Visok | **Delimično** — Pers:157-158 potvrđeno, ali legalne tranzicije ne vode nazad u SENDING posle Da-stanja (Val:24-29,37-41) pa se postojeće „Da“ ne briše; `SEFSentAt` se ionako ne prepisuje (Pers:153-155); ostaje semantička rupa za unknown-outcome pokušaje | P3 | Opciono `EverSubmitted` polje | S |
| 8 | Nema unique `FakturaID+VersionNo` | Visok | **Tačno** mehanički (Pers:217-295 bez provere), ali single-writer + TX čine kolziju logičkim bugom, ne race-om | Prihvaćeno | Audit provera po potrebi | S |
| 9 | Nema CAS | Visok | **Dizajnersko ograničenje** — single-writer | Prihvaćeno | — | — |
| 10 | ID max+1 bez rezervacije | Visok | **Dizajnersko ograničenje** — `GetNextID` (Pers:252,371) je standardni obrazac celog projekta; single-writer | Prihvaćeno | — | — |
| 11 | Pozicioni insert (submission 20 kolona, event 9) | Visok | **Tačno** — Pers:258-282 i Pers:377-390; `Require*Schema` proverava postojanje, NE redosled (Pers:716-749); direktno suprotno naučenom pravilu projekta (CLAUDE.md §4) — insert kolone tiho korumpira audit/retry izvor | P2 | Upis po imenu kolone ili order-assert u schema check | M |
| 12 | Empty SubmissionID u save-result je silent no-op | Visok | **Tačno** — Pers:305 `Exit Sub`; orchestration danas garantuje neprazan ID, ali bug bi bio nevidljiv | P3 | `Err.Raise` umesto `Exit Sub` | S |
| 13 | Hash nije obavezan | Visok | **Tačno** — create validira sve osim `payloadHash` (Pers:229-247); pozivaoci ga uvek računaju (Svc:53,66) | P3 | Require hash u create | S |
| 14 | Orphan submission/event (nema FK) | Visok | **Tačno** — nema provere postojanja fakture; pozivaoci validiraju pre | P3 | Provera postojanja u create | S |
| 15 | Stornirani submission utiče na state | Visok | **Delimično** — filter samo po FakturaID (Pers:523) potvrđen, ali **ništa danas ne postavlja** `Stornirano="Da"` na submission (create piše „Ne“, Pers:278; nema writera) — mrtvo polje | P3 | Filtrirati `Stornirano` u `HasSuccessful`/`GetLast` | S |
| 16 | DocumentID Text format je best-effort | Visok | **Delimično** — `On Error Resume Next` potvrđen (Pers:450-462; i Pers:337-345), ali defanzivni read `Format$(v,"0")` (Pers:775-780) pokriva legacy Double; SEF ID dužine su ispod granice preciznosti Double-a | P3 | Postcondition provera (pročitaj nazad i uporedi) | S |
| 17 | `SubmittedAt`=`FinishedAt` | Visok | **Delimično** — Pers:329-330 oba `Now`, ali `CreatedAt` (Pers:264) nastaje pre HTTP-a pa je grubo trajanje izvedivo | P3 | `SubmittedAt` upisivati u TX1 fazi | S |
| 18 | LastSyncAt je dvosmislen | Srednji | **Tačno** — Pers:161-164 (i na SYNC_ERROR) + Pers:465-487 | P3 | Razdvojiti attempt/success | S |
| 19 | Clear pointer nema event | Srednji | **Delimično** — helper ćuti (Pers:677-699), ali jedini pozivalac `PrepareRejectedInvoiceForResubmit` loguje event u istoj TX (Val:414-421) | P3 | Event u helperu radi budućih pozivalaca | S |
| 20 | Slobodan workflow/format tekst | Srednji | **Tačno** — bez enum/whitelist provere (Pers:237-247); pozivaoci koriste konstante | P3 | `Select Case` whitelist | S |
| 21 | Raw payload retention | Srednji | **Dizajnersko ograničenje** — isto kao FM-0037 #25 | Prihvaćeno | — | — |
| 22 | Event nema OperationID/CausationID | Srednji | **Tačno** — 9-kolona event šema (Pers:377-390, 739-749) | P3 | Nova kolona kroz `Ensure*Schema` + prosleđivanje run ID-ja | M |
| 23 | Naziv `Persistance` je typo | Nizak | **Tačno** — Pers:1 vs sopstveni komentar Pers:5 („modSEFPersistence“) | Prihvaćeno (rename = churn na mnogo call-site-ova bez funkcionalne koristi) | — | — |

**Bilans:** 23/23 provereno; 13 Tačno, 7 Delimično, 3 Dizajnersko ograničenje, 0 Netačno. Hitnost: P1×0, **P2×3** (#1 netačan poreski/audit indikator za REJECTED, #6 fail-closed idempotency read, #11 pozicioni insert protivan pravilu projekta), P3×15, Prihvaćeno×5. Tri od šest „Kritičan“ ocena FM-a su precenjene (fail-late umesto korupcije: #3; TX-atomičnost drži pointer: #2; prazna verzija je legitiman prvi put: #4).

**Ukupno: 78/78 rizik-redova verifikovano (28+27+23), svi Kritičan/Visok sa citatima file:line.** Najviši prioriteti preko sva tri fajla: FM-0037 #1 (P1, frm poruka „Faktura poslata“ za odbijenu fakturu), FM-0037 #6 (P1, javni `Test_Cancel/Storno` makroi sa pravnim side effectom), zatim P2 grupa: FM-0037 #13 + FM-0038 #2/#3 (recovery/refresh result contract i fail-open statusi), FM-0037 #5/#11 (cancel/storno intent + neiskorišćeno `SEF_STORNO` stanje), FM-0037 #2 (potvrditi SEF requestId dedup na demo), FM-0039 #1/#6/#11.

---

## Delta blok 3 — SEF validator i mapper (FM-0040…FM-0041, 64 stavke) [sidro f6313dc]

Sve provere su završene. Imam kompletnu sliku za svih 64 stavke. Finalni izveštaj:

---

# Audit rizik-nalaza FM-0040 i FM-0041 (commit `f6313dc`)

**Skraćenice fajlova** (svi u `/tmp/claude-0/.../scratchpad/wt-f6313dc/src-vba/`): Val=`modSEFValidator.bas`, Map=`modSEFMapper.bas`, Pers=`modSEFPersistance.bas`, Svc=`modSEFService.bas`, Sync=`modSEFStatusSync.bas`, Sto=`modStorno.bas`, Nov=`modNovac.bas`, Tax=`modSEFTax.bas`, Cfg=`modConfig.bas`, Cli=`modSEFClient.bas`.

**Ključne unakrsne činjenice** (utiču na više redova): (1) `GetSingleRowIndexByKey` (Pers:408-433) diže grešku na duplikat i poziva se u SVIM write putanjama **pre** HTTP-a (Svc:128 pre Svc:167) → first-match u read sloju je fail-closed za tok slanja. (2) `WF_SEF_UNKNOWN` niko nikad ne upisuje (0 pogodaka van matrice/testova). (3) `ValidateSEFPayload`, `IsFinalSEFStatus`, `IsPendingSEFStatus`, `GetSEFDisplayStatus` nemaju nijednog produkcionog pozivaoca. (4) `GetDefaultTaxPercent`=10 je hardkodovana konstanta (Tax:2-3), ne config. (5) XmlAmount je locale-bezbedan (`Format$ "0.00"` + `Replace ","→"."`, Map:627-635); dijakritika je bezbedna jer WinHttp BSTR šalje kao UTF-8 (Cli:33,35). (6) Model IMA stavka-level `Stornirano` (Sto:1024-1052) i `OsirocenoOd` (Sto:581-582,1095-1119). (7) `ApplyAvansToFaktura` se automatski poziva pri kreiranju fakture (modFaktura.bas:331 → Nov:525+).

### FM-0040 — `modSEFValidator.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | `SEF_UNKNOWN` je dead-end | Kritičan | **Delimično** — matrica dozvoljava ulaz (Val:19) bez izlaza (nema `Case WF_SEF_UNKNOWN`, Val:46-48 diže grešku); ALI nijedan kod ne upisuje to stanje → danas nedostižno | P3 | Dodati `Case WF_SEF_UNKNOWN` (izlaz kroz reconciliation ka SENT/TECH_FAILED/ACCEPTED/REJECTED) ili ga izbaciti iz Val:19 dok nema writera | S |
| 2 | Faktura validation first-match | Kritičan | **Delimično** — validator staje na prvom pogotku (Val:95-104); ali duplikat obara slanje fail-closed pre HTTP-a (Svc:128 → Pers:427-430 `ERR_SEF_DUPLICATE`) | P3 | I u validatoru koristiti postojeći exact-single obrazac (`GetSingleRowIndexByKey`) | S |
| 3 | Lokalni storno nije proveren | Kritičan | **Tačno** — validator ne čita `Stornirano` (Val:130-154); `StornoFaktura` ne menja SEF workflow (Sto:696-719); frmSEF combo bez filtra (frmSEF.frm:264-276) → stornirana faktura je end-to-end sendable | P1 | U `ValidateFakturaForSEF` raise ako `Stornirano="DA"`; u `LoadFaktureIntoCombo` filtrirati stornirane | S |
| 4 | Successful-submission guard fail-soft | Kritičan | **Tačno** — `GetSEFSubmissionsForFaktura` u EH vraća `Empty` (Pers:528-530) → `HasSuccessfulSEFSubmission`=False (Pers:570-572) → kvar audit read-a dozvoljava novo slanje (Val:156-159 prolazi) | P1 | U EH `GetSEFSubmissionsForFaktura` re-raise (guard mora biti fail-closed) | S |
| 5 | Resubmit ne potvrđuje korekciju | Kritičan | **Tačno** — `PrepareRejectedInvoiceForResubmit` proverava samo `currentState=SEF_REJECTED` (Val:394-399), nikakav dokaz izmene | P2 | Pre prelaza uporediti novi (rebuild) payload hash sa hash-om odbijenog submissiona; zahtevati razliku ili eksplicitnu potvrdu + reason | M |
| 6 | Resubmit upisuje `SEFStatus=READY` | Kritičan | **Tačno** — Val:406-412 (`sefStatus:=WF_SEF_READY`) upisuje lokalni state u kolonu remote statusa, suprotno dokumentovanom modelu (Pers:83-86); remote REJECTED se gubi iz kolone (ostaje u tblSEFSubmission/event logu) | P2 | Proslediti `sefStatus:=""` (zadržava se REJECTED jer Pers:125-127 ne piše prazno) | S |
| 7 | Stari DocumentID ostaje kroz novu generaciju | Kritičan | **Tačno** — resubmit ne čisti `SEFDocumentId` (Val:406-414; Pers:129-132 prazan param = zadrži staro); recovery zaglavljenog SENDING preferira refresh po starom ID-ju (Svc:652-653) → nova generacija dobija status starog odbijenog dokumenta | P1 | Pri resubmit prepare arhivirati stari ID u event details pa obrisati kolonu (direktan `RequireUpdateCell` u istom TX) | M |
| 8 | Stavke se proveravaju samo po postojanju | Visok | **Tačno** — `ValidateFakturaHasStavke` samo `FindRows.count>0` (Val:184-190); ne filtrira `Stornirano` ni `OsirocenoOd`, a obe kolone postoje u modelu (Sto:1042, Sto:1112) | P2 | U proveru uključiti samo aktivne stavke (Stornirano≠DA, OsirocenoOd prazan) | S |
| 9 | Nema header/detail reconciliation | Visok | **Delimično** — validator ga nema (Val:58-170), ali send tok odmah zatim radi troslojni reconciliation u `BuildSEFInvoiceDto` (Map:220-222) → nekonzistentan total ne prolazi do slanja | P3 | Ništa hitno; po želji izvući zajednički reconciliation helper | S |
| 10 | Kupac lookup first-match | Visok | **Tačno** — dva odvojena `LookupValue` (Val:236-237), bez exact-single; duplikat KupacID prolazi neotkriven (oba čitanja konzistentno vraćaju prvi red — hibrid praktično isključen u single-writer režimu) | P3 | `FindRows(TBL_KUPCI,"KupacID").count=1` provera | S |
| 11 | PIB je samo nonempty | Visok | **Tačno** — Val:244-247 samo neprazan; bez formata/dužine/checksum-a; nevalidan PIB → sigurna remote rejekcija ili pogrešan identitet (Map ga i prefiksira sa "RS") | P2 | Provera: 9 cifara + mod-11 kontrolna cifra (jedan mali helper, koristiti i u mapperu) | S |
| 12 | Seller podaci nisu validirani | Visok | **Delimično** — validator ih ne proverava, ali mapper fail-fast za SELLER_NAME/PIB pre slanja (Map:108-114); adresa/MB/račun zaista ostaju neproverni (→ FM-0041 #17) | P2 | Rešava se kroz FM-0041 #17 (obavezna seller config polja u serializeru) | S |
| 13 | Payload validacija je substring test | Visok | **Tačno** kao opis (Val:204-212); napomena: funkcija nema produkcionog pozivaoca (samo modSEFTests) — realni tok koristi `ValidateGeneratedUBL` u mapperu | P3 | Ili je uključiti u send tok sa DOM parse-om, ili uklonati/označiti kao legacy | S |
| 14 | Payload identitet nije vezan za fakturu | Visok | **Tačno** — ni Val:204-212 ni `ValidateGeneratedUBL` (Map:963-979) ne porede `<cbc:ID>` sa `BrojFakture` | P3 | U `ValidateGeneratedUBL` dodati `InStr(xml, "<cbc:ID>" & broj & "</cbc:ID>")` provere | S |
| 15 | Cancel/storno ne proveravaju local workflow/generation | Visok | **Tačno** — Val:316-380 traže samo DocumentID + status tekst iz kolone (može biti stale — nema prisilnog refresh-a pre akcije) | P2 | Pre cancel/storno pozvati `RefreshSEFStatus_TX` pa tek onda validirati status | S/M |
| 16 | Storno dozvoljen za REJECTED | Visok | **Tačno** kao kod-činjenica (Val:367); da li je API-korektno — nije proverivo statički (SEF ugovor) | P3 | Potvrditi uz SEF API dokumentaciju; ako nije podržano, izbaciti REJECTED iz whitelist-e | S |
| 17 | Status/DocumentID read nije jedan snapshot | Visok | **Delimično** — jesu dva odvojena čitanja (Val:324, Val:331), ali oba first-match po istom redosledu → isti red; hibrid samo uz izmenu tabele između poziva (single-writer desktop) | P3 | Jedno čitanje reda (row index + oba polja) | S |
| 18 | Last submission pointer se briše bez lineage podataka | Visok | **Delimično** — pointer se briše (Val:414, Pers:692), ali kompletna istorija ostaje u tblSEFSubmission (FakturaID+VersionNo+CreatedAt) → lineage rekonstruktibilan, samo nije eksplicitan | P3 | U resubmit event details upisati stari SubmissionID/DocumentID/hash | S |
| 19 | Final/pending status liste driftuju | Visok | **Delimično** — drift potvrđen: Val:454 nema `"CANCELED"`, Sync:118 i Sync:454 imaju; ALI `IsFinal/IsPendingSEFStatus` nemaju nijednog pozivaoca → bez efekta danas | P3 | Konsolidovati u jednu javnu klasifikaciju (Sync verzija kao izvor istine) | S |
| 20 | Display status skriva lokalni problem | Visok | **Delimično** — logika potvrđena (Val:475-479, SEFStatus apsolutni prioritet) i lokalna stanja se zaista upisuju u SEFStatus (Svc:84,131,674; Sync:168); ali helper nema pozivaoce, a FM primer SYNC_ERROR/SENT ne nastaje (pri SYNC_ERROR se i SEFStatus prepiše, Sync:168) | P3 | Kad helper uđe u upotrebu: kombinovani prikaz workflow+status+sync health | S |
| 21 | Transitioni nisu normalizovani | Srednji | **Tačno** — Val:4-56 bez Trim/UCase; izloženost mala (ulazi su konstante, a read je trimovan — Pers:779), rizik samo ručni unos u ćeliju | P3 | `UCase$(Trim$())` na oba parametra na ulazu | S |
| 22 | Self-transition nema no-op semantiku | Srednji | **Tačno** — svaki state→isti state pada; sync sloj to svesno zaobilazi refresh-only granom (Sync:385-393), direktan caller dobija grešku | P3 | No-op grana (oldState=newState → Exit Sub) ili poseban povratni kod | S |
| 23 | HTTPS validacija je površna | Srednji | **Tačno** — samo prefiks provera (Val:277-280); `https://` bez hosta prolazi | P3 | Minimalno: zahtevati host deo; opciono allowlist demo/prod hostova + env flag | S |
| 24 | API key placeholder prolazi | Srednji | **Tačno** — samo neprazan (Val:272-275) | P3 | Minimalna dužina + odbaciti očigledne placeholdere | S |
| 25 | Error poruka hardkoduje config tabelu | Srednji | **Netačno** — `GetConfigValue` ČITA upravo tblSEFConfig (Cfg:786), poruka Val:269/274 je ispravna. (Obrnuta greška postoji u mapperu: „missing in tblConfig" Map:109/113 — tamo je pogrešno) | P3 | Ispraviti poruke u Map:109/113 na tblSEFConfig | S |
| 26 | Resubmit event nema correction detalje | Srednji | **Tačno** — details samo `PreviousState=` (Val:416-421) | P3 | Dodati stari SubmissionID/DocumentID/hash/verziju (isti fix kao #18) | S |
| 27 | Rollback rezultat nije poznat | Srednji | **Tačno** — rollback pod `On Error Resume Next`, ishod se ne vraća caller-u (Val:436-448) | P3 | Uhvatiti rollback grešku i dopisati je u rethrow poruku | S |
| 28 | Status classification je duplicirana | Srednji | **Tačno** — tri liste: Val:451-471 vs Sync:101/118 vs Sync:441-459 | P3 | Jedna javna klasifikacija (uz #19) | S |
| 29 | Nema auth/approval/period policy-ja | Srednji | **Dizajnersko ograničenje** — tačna činjenica, ali aplikacija je single-user desktop bez modela rola; policy nema gde da se osloni | Prihvaćeno | Zabeležiti kao svesno ograničenje; eventualno confirm-dijalog za storno/cancel u UI | M |
| 30 | Naziv `StorniranoOnSEF` neujednačen | Nizak | **Tačno** (Val:349) | Prihvaćeno | Ne dirati (rename lomi pozivaoce bez dobiti); ujednačiti pri sledećem većem refaktoru | S |
| 31 | Mešani engleski/srpski error tekstovi | Nizak | **Tačno** — Err poruke engleske, UI katalog (`modPoruke`) srpski | Prihvaćeno | Tehničke poruke ostaviti EN; korisničke prikaze rutirati kroz `Poruka()` | M |

**Bilans FM-0040:** 31 stavka — **Tačno 21**, **Delimično 8** (#1, #2, #9, #12, #17, #18, #19, #20), **Netačno 1** (#25), **Dizajnersko ograničenje 1** (#29). Hitnost: **P1 ×3** (#3 stornirana faktura sendable, #4 fail-soft idempotency guard, #7 stale DocumentID kroz resubmit), **P2 ×5**, **P3 ×20**, **Prihvaćeno ×3**. Nema P0 — nijedan nalaz sam po sebi ne proizvodi malformed/pogrešnu fakturu u normalnom toku. Ključna kalibracija: first-match nalazi (7 „Kritičan/Visok" redova) su ublaženi time što svaka write putanja ide kroz strict `GetSingleRowIndexByKey` pre HTTP-a, a mrtvi helperi (`ValidateSEFPayload`, display/final/pending) nemaju produkcione pozivaoce. Najisplativiji paket: #3+#4+#6+#8 su svi S-napor, jedan mali PR.

### FM-0041 — `modSEFMapper.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | Quantity/price se skraćuju na 2 decimale | Kritičan | **Tačno** — line net iz pune preciznosti (Map:190 `Round(qty*price,2)`), a `InvoicedQuantity`/`PriceAmount` idu kroz isti `XmlAmount` sa 2 dec (Map:561, Map:577, Map:627-635); qty=1,234 → XML qty„1.23"×100,00=123,00 ≠ LineExtensionAmount 123,40 | P1 | Odvojeni formatteri: `XmlQuantity` (3+ dec) i `XmlUnitPrice` (pun broj decimala); plus provera skale ulaza | S |
| 2 | `PrepaidAmount=0` i puni PayableAmount | Kritičan | **Tačno** — hardkod `0.00` + `PayableAmount=TotalGross` (Map:547, Map:549), a avans kupca se automatski vezuje za fakturu već pri kreiranju (modFaktura.bas:331 → Nov:525+) | P2 | Poslovna odluka o modelu avansa na SEF-u (avansni račun tok ne postoji); minimalno: čitati vezane avanse iz tblNovac i popuniti Prepaid/Payable | M |
| 3 | DTO nije stvarni snapshot | Kritičan | **Tačno** — `clsSEFInvoiceSnapshot` nema adrese/tax/payment/note (potvrđen spisak polja), serializer ponovo čita config+tblKupci+`Date` (Map:355-389, Map:409); ublaženo time što se poslati XML čuva i retry ga reuse-uje (Svc:35-54) | P2 | Proširiti DTO na sva emitovana polja; serializer učiniti čistom funkcijom DTO→XML | L |
| 4 | Tax rate se čita ponovo (stopa A vs B) | Kritičan | **Netačno** (danas) — oba čitanja (Map:42, Map:387) vraćaju hardkodovanu konstantu 10 (Tax:2-3, nije config) → divergencija nemoguća; rizik postaje realan tek ako Tax pređe na config | P3 | Preventivno: percent+categoryID upisati u DTO pri build-u i emitovati iz DTO-a (deo #3) | S |
| 5 | Force-today IssueDate, DueDate od starog datuma | Kritičan | **Tačno** — `dueDate=DateAdd(...,dto.InvoiceDate)` (Map:372) pre grane koja IssueDate zamenjuje sa `Date` (Map:411-412) → za stariju fakturu DueDate < IssueDate; nema validacije odnosa | P1 | DueDate računati od stvarno emitovanog IssueDate + raise ako DueDate<IssueDate | S |
| 6 | First-match FakturaID | Kritičan | **Delimično** — jeste first-match (Map:48-75), ali duplikat obara send pre HTTP-a u persistence sloju (Svc:128 → Pers:427-430) | P3 | Exact-single provera i u build-u (isti obrazac kao Pers) | S |
| 7 | Stornirane/superseded stavke nisu filtrirane | Kritičan | **Tačno** — petlja uzima svaki red sa FakturaID (Map:158-214) bez `Stornirano`/`OsirocenoOd` filtra; realan put: storno prijemnice posle fakturisanja ostavlja fakturu aktivnom sa orphaned stavkama (Sto:581-582, Sto:1095-1119) i one ulaze u UBL | P2 | U build petlji raise ako je stavka `Stornirano="DA"` ili `OsirocenoOd`≠"" | S |
| 8 | Slab 31-bit payload hash | Kritičan | **Delimično** — algoritam i `Asc` code-page zavisnost potvrđeni (Map:747-768, Map:757); ALI hash se koristi samo kao audit/correlation trag (Svc:53,66) — nigde kao idempotency/integrity ključ (`ShouldReuseLastSubmission` ga ne poredi, Svc:605-625) | P3 | Preći na SHA-256 nad UTF-8 bajtovima tek kad hash dobije autoritativnu ulogu; do tada dokumentovati da je samo trag | M |
| 9 | Class se gubi u stvarnom UBL-u | Visok | **Tačno** — klasa ulazi u DTO/JSON (Map:201, Map:283) ali je UBL serializer ne emituje (Map:559-580); linije I i II klase iste prijemnice postaju istoimene sa različitim cenama | P2 | Dodati klasu u `cbc:Name` (najmanja delta) ili `AdditionalItemProperty` | S |
| 10 | Buyer/seller podaci su current master/config | Visok | **Tačno** — build čita trenutni tblKupci/config (Map:94-95, Map:105-106), serializer još jednom (Map:355-379); promena mastera menja payload sledeće serijalizacije | P2 | Rešava se sa #3 (pun snapshot u DTO) | M |
| 11 | DeliveryDate je current Prijemnica datum | Visok | **Tačno** — max `Datum` iz tblPrijemnica u trenutku slanja (Map:646-717), ne snapshot sa fakture | P3 | Snapshot datuma prometa pri fakturisanju (deo #3); izmena prijemnice posle fakturisanja je ionako redak/kontrolisan tok | S |
| 12 | Duplicate PrijemnicaID first-match | Visok | **Tačno** — `LookupValue` prvi red (Map:180-181, Map:689), bez exact-single (koji u storno sloju postoji — Sto:1055-1075) | P3 | `FindRows.count=1` provera po prijemnici u build petlji | S |
| 13 | No exact buyer/seller snapshot | Visok | **Tačno** — isti koren kao #3/#10 (dupliran red u FM tabeli); buyer name/PIB u DTO, ostalo živo pri serijalizaciji | P2 | Isti fix kao #3 | M |
| 14 | Header/line VAT rounding politika nije skalabilna | Visok | **Tačno** — header VAT `Round(net×10%,2)` (Map:68) vs zbir line VAT-ova (Map:191,209), tolerancija fiksna 0,02 (Map:986); od ~5+ linija worst-case prelazi toleranciju → legitimna faktura blokirana (fail-closed, ne pogrešna) | P2 | Jedna politika: header VAT = Σ(line VAT) pri kreiranju fakture; ili toleranciju vezati za broj linija | M |
| 15 | Mutable DTO serializer ne re-reconciliuje | Visok | **Tačno** — sva polja `Public` (potvrđeno u obe klase), `ValidateSEFDtoForUBL` ne proverava neto=qty×cena niti Σlinija=total (Map:833-941); jedini realni pozivalac je service sa svežim build-om | P3 | Pozvati `ValidateSEFTotalMatch` i u serializeru | S |
| 16 | Generated UBL validacija je substring smoke check | Visok | **Tačno** — tri `InStr` markera (Map:963-979); ovo JE jedina automatska provera stvarnog outbound XML-a (`ValidateSEFPayload` se ne poziva u toku) | P2 | `MSXML2.DOMDocument.LoadXML` + `parseError` well-formedness check (bez novih zavisnosti, MSXML je sistemski) | S |
| 17 | Empty required seller/buyer/account elementi | Visok | **Tačno** — seller street/city/postal/MB (Map:439-452) i račun (Map:519) emituju se bezuslovno i prazni; buyer street/city/MB (Map:476-493) takođe; DTO validacija ih ne pokriva → prazan config = remote rejection umesto lokalne greške | P2 | Upotrebiti postojeći (mrtav!) `GetRequiredSEFConfig` (Map:943) za obavezna seller polja; buyer polja emitovati uslovno | S |
| 18 | Foreign buyer se tretira kao RS PIB | Visok | **Dizajnersko ograničenje** — kod potvrđen (default RS Map:381, schemeID 9948 Map:469, `"RS"&PIB` Map:487), ali otkup voća ima domaće kupce; nema stranog scenarija | P3 | Eksplicitni guard: raise ako `Drzava`≠RS (pretvara tiho pogrešan payload u jasnu grešku) | S |
| 19 | Jedna tax kategorija za sve stavke | Visok | **Dizajnersko ograničenje** — potvrđeno (jedan TaxSubtotal Map:526-537, ista kategorija po liniji Map:569-573); domen je jednoobrazan (voće, 10% S) | Prihvaćeno | Zabeležiti ograničenje; per-line model tek uz stvarnu potrebu | L |
| 20 | Currency hardkodovan RSD | Visok | **Dizajnersko ograničenje** — Map:116; lokalni faktura model nema valutu uopšte | Prihvaćeno | Ništa; multi-currency je novi domen zahteva | L |
| 21 | InvoiceTypeCode samo 380 | Visok | **Dizajnersko ograničenje** — Map:418; storno ide preko SEF storno API-ja (Svc:526+), ne kroz kreditni dokument | Prihvaćeno | Zabeležiti; credit note tek ako poslovni tok to zatraži | L |
| 22 | XML control chars nisu validirani | Visok | **Tačno** — `XmlEscape` samo 5 entiteta (Map:612-625), bez uklanjanja nelegalnih kontrolnih znakova; verovatnoća niska (unos iz Excel ćelija; LF je legalan u XML-u), dijakritika bezbedna (WinHttp→UTF-8, Cli:33-35) | P3 | U `XmlEscape` strip `[x00-x08 x0B x0C x0E-x1F]` | S |
| 23 | Line order zavisi od fizičkog reda | Visok | **Tačno** — redosled = fizički red tabele, ID=`CStr(i)` (Map:158, Map:560); posledica ograničena: hash nije ključ, retry reuse-uje sačuvani XML | P3 | Sortirati linije po StavkaID pre serijalizacije | S |
| 24 | Broj prijemnice nije obavezan | Srednji | **Tačno** — Map:163 bez provere; napomena: fallback „Roba po prijemnici" (Map:186-188) je mrtav kod jer opis uvek sadrži literal „po prijemnici" (Map:184) pa nikad nije prazan | P3 | Zahtevati neprazan `BrojPrijemnice` u build petlji | S |
| 25 | Klasa nije validirana | Srednji | **Tačno** — raw copy (Map:201), bez kanonskog skupa | P3 | Provera protiv dozvoljenog skupa (I/II) pri build-u | S |
| 26 | Seller account nije obavezan | Srednji | **Tačno** — Map:360 + bezuslovni prazan `<cbc:ID>` (Map:518-520) | P2 | Deo #17 (obavezan `SELLER_ACCOUNT` kroz `GetRequiredSEFConfig`) | S |
| 27 | Payment/note/category su current config | Srednji | **Tačno** za payment means/note/period code (Map:363-364, Map:389, defaulti Map:367-368); tax category je zapravo hardkod (Tax:6-7), ne config | P3 | U DTO uz #3; usput ispraviti poruke „tblConfig" (Map:109,113) | S |
| 28 | Due-days nema gornju granicu | Srednji | **Tačno** — samo `<0` blokiran (Map:791-794); ogroman Long → `DateAdd` greška ili besmislen datum | P3 | Cap (npr. 365) u `GetSEFPaymentDueDays` | S |
| 29 | `DA` config nije trim/canonical | Srednji | **Tačno** — `UCase$` bez `Trim$` (Map:409-411); napomena: `GetConfigValue` sam trimuje (Cfg:791), pa „DA " ipak radi — ne radi „da/true/1" iz drugih izvora; rizik je fail-safe smer (flag se ne aktivira) | P3 | `UCase$(Trim$())` + centralni bool-parser za DA/NE | S |
| 30 | Ručna string XML gradnja | Srednji | **Tačno** — ceo serializer je konkatenacija (Map:391-584); rizik održavanja realan, ali pristup je svesno bez zavisnosti | P3 | Zadržati string gradnju, ali dodati DOM parse validaciju (#16) kao sigurnosnu mrežu | M |
| 31 | Public debug makroi | Srednji | **Tačno** — 5 javnih Test/Debug procedura u produkcionom modulu, hardkodovan `FAK-00001`, `Debug.Print` buyer/finansijskih podataka (Map:1001-1139); ne šalju HTTP | P3 | Premestiti u `modSEFTests` (već postoji!) | S |
| 32 | Debug JSON nije potpuni JSON | Srednji | **Tačno** — escape samo `\`, `"`, CR/LF (Map:302-313), bez `\t` i U+0000-1F; nije transport (komentar Map:232-233) | P3 | Dopuniti escape petljom za <32; ili prihvatiti kao debug-only | S |
| 33 | Mrtvi/duplirani helperi | Nizak | **Tačno** — `GetFakturaIssueDate` (Map:804-829) i `GetRequiredSEFConfig` (Map:943-961) imaju 0 pozivalaca (proveren ceo src-vba) | P3 | `GetRequiredSEFConfig` upotrebiti za #17/#26; `GetFakturaIssueDate` obrisati (duplira Map:719-745) | S |

**Bilans FM-0041:** 33 stavke — **Tačno 26**, **Delimično 2** (#6, #8), **Netačno 1** (#4 — tax stopa je hardkodovana konstanta, divergencija build/serialize danas nemoguća), **Dizajnersko ograničenje 4** (#18-21). Hitnost: **P1 ×2** (#1 qty/price precision — aritmetički nekonzistentan pravni dokument čim podatak ima >2 decimale; #5 DueDate pre IssueDate uz force-today), **P2 ×10**, **P3 ×18**, **Prihvaćeno ×3**. Nema P0: XML escaping i decimalno formatiranje su za normalne podatke ispravni (locale-bezbedan `XmlAmount`, WinHttp šalje UTF-8), pa malformed XML nastaje samo uz egzotičan ulaz. Posebno vredno: FM-ov predlog #17 se rešava aktiviranjem već postojećeg mrtvog helpera (#33). Najisplativiji paket: #1+#5+#7+#9+#16+#17 — svi S-napor, direktno smanjuju šansu odbijene/pogrešne fakture.

**Ukupno oba fajla:** 64/64 stavke provereno protiv koda; 47 Tačno, 10 Delimično, 2 Netačno, 5 Dizajnersko ograničenje. P1 ukupno 5 (Val: stornirana-sendable, fail-soft guard, stale DocumentID; Map: qty/price precision, force-today DueDate) — svi sa S/M minimal-delta popravkama.

---

## Delta blok 4 — SEF DTO klase, klijent, tax (FM-0042…FM-0046, 77 stavki) [sidro f6313dc]

Sve provere su završene (klijent, mapper/serializer, service, status-sync, persistence, validator, tax). Kompletan audit svih 77 stavki sledi.

# Audit rizik-nalaza FM-0042..FM-0046 (SEF sloj) — kod na `f6313dc`

Ključne verifikacione tačke: DTO klase su potvrđeni field-bagovi (`clsSEFInvoiceSnapshot.cls:12-31`, `clsSEFLine.cls:12-20`, `clsSEFResponse.cls:12-22`); jedini writer/consumer lanac je `modSEFMapper` → `modSEFService` → `modSEFClient` → `modSEFStatusSync`/`modSEFPersistance` (single-writer kalibracija primenjena).

### FM-0042 — `clsSEFInvoiceSnapshot.cls`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | Snapshot potpuno mutable | Kritičan | **Tačno** (sva polja `Public`, cls:12-31; nema seal) | P3 | Nijedan postojeći caller ne mutira između build i serialize (modSEFService.bas:110-167 sekvencijalno); seal tek uz budući refactor | L |
| 2 | Public mutable `lines` | Kritičan | **Tačno** (cls:31, nema AddLine/copy) | P3 | Isti paket kao #1; odmah ništa | L |
| 3 | Mutable line objekti | Kritičan | **Tačno** (Collection drži reference na `clsSEFLine`) | P3 | Isti paket kao #1 | L |
| 4 | Serializer ne ponavlja reconciliation | Kritičan | **Tačno** — `ValidateSEFDtoForUBL` (modSEFMapper.bas:833-941) proverava samo ≥0/ne-prazno; `ValidateSEFTotalMatch` samo u build-u (modSEFMapper.bas:220-222) | **P2** | Pozvati `ValidateSEFTotalMatch` nad zbirom `dto.lines` i u `ValidateSEFDtoForUBL` (postojeći helper, ~6 linija) | S |
| 5 | DTO nije kompletan UBL snapshot | Kritičan | **Delimično** — serializer zaista re-čita config/master (modSEFMapper.bas:355-364, 374-379), ali build+serialize su jedan sinhroni tok, a retry šalje ARHIVIRANI XML (modSEFService.bas:38), ne re-serializaciju | P3 | Bez izmene; arhivirani request body već daje reproducibilnost | M |
| 6 | Nema tax/payment snapshot-a | Visok | **Delimično** — tačno kao odsustvo (samo `TotalVat`), ublaženo istim jednoprocesnim tokom + arhivom XML-a | P3 | Rešiti kroz FM-0046 selidbu stope u config | M |
| 7 | Nema precision metadata | Visok | **Delimično** — klasa je samo `Double` skladište; stvarni problem je `XmlAmount` 2-dec za qty/cenu (modSEFMapper.bas:561,577,627-636) vs pun precision u neto (190) | **P2** | U mapperu poseban `XmlQty`/`XmlPrice` format (3-4 dec) — fix pripada FM-0043 #4 | S |
| 8 | Totali nezavisno mutable | Visok | **Tačno** (cls:27-29; identitet gross=net+vat niko ne re-proverava posle build-a) | P3 | Pokriveno predlogom #4 | S |
| 9 | Buyer može biti hibridan | Visok | **Delimično** — naziv/PIB iz DTO, adresa re-lookup po BuyerID (modSEFMapper.bas:374-379); u istom procesu, single-writer → prakt. nedostižno | P3 | Bez izmene | M |
| 10 | Nema provenance/version tokena | Visok | **Delimično** — postoje payload hash (`ComputePayloadHash`), arhiviran request body i submission red | P3 | Bez izmene | M |
| 11 | `lines` nije inicijalizovan | Srednji | **Tačno** (nema `Class_Initialize`; mapper ga postavlja, modSEFMapper.bas:119) | P3 | Dodati `Class_Initialize` sa `Set lines = New Collection` | S |
| 12 | Validacija izvan klase | Srednji | **Tačno** (u modSEFMapper.bas:833) | P3 | Prihvatljivo za VBA DTO obrazac | M |
| 13 | Currency slobodan string | Srednji | **Tačno** — ali mapper hardkoduje "RSD" (modSEFMapper.bas:116); jedini writer | P3 | Bez izmene | S |
| 14 | Field bag, naziv obećava više | Nizak | **Tačno** | Prihvaćeno | Event. preimenovanje u `Dto` uz refactor | S |

**Bilans:** 14/14 provereno · Tačno 9 · Delimično 5 · P2×2, P3×11, Prihvaćeno×1. Nijedan „Kritičan" nema aktivan put eksploatacije u tekućem kodu; jedini vredan brzi dobitak je re-check totala u serializeru (#4).

### FM-0043 — `clsSEFLine.cls`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | Sva polja mutable | Kritičan | **Tačno** (cls:12-20) | P3 | Bez mutirajućih callera; seal uz budući refactor | L |
| 2 | Finansijske vrednosti nisu derived | Kritičan | **Tačno** — računa ih samo mapper (modSEFMapper.bas:190-192); `ValidateSEFDtoForUBL` proverava samo ≥0 (926-939) | **P2** | U `ValidateSEFDtoForUBL` re-proveriti `neto=Round(kolicina*cena,2)` i `iznos=neto+pdv` po liniji | S |
| 3 | Serializer nema seal dokaz | Kritičan | **Tačno** kao odsustvo | P3 | Pokriveno #2 (aritmetički re-check je praktičan „dokaz") | S |
| 4 | Nema precision/scale modela → UBL mismatch | Visok | **Delimično** — „direktan uzrok" je `XmlAmount` (2 dec) u mapperu (modSEFMapper.bas:561,577), ne klasa; `Double` čuva pun precision | **P2** | Qty/cena sa više decimala u UBL (poseban format helper); regres: qty 3 dec | S |
| 5 | Nema tax policy na liniji | Visok | **Tačno** (samo `pdv` iznos, cls:19); stopa se re-čita u serializeru (modSEFMapper.bas:387) | P3 | Rešava se FM-0046 paketom | M |
| 6 | Nema UOM/currency, KGM hardkodovan | Visok | **Dizajnersko ograničenje** — sav promet je otkup voća u kg/RSD; KGM: modSEFMapper.bas:561 | P3 | Bez izmene dok domen ne dobije drugi UOM | M |
| 7 | Nema stable line ID | Visok | **Tačno** — UBL ID = indeks kolekcije (modSEFMapper.bas:560); redosled determinističan iz tabele | P3 | Uz refactor: FakturaStavka sequence | M |
| 8 | Source identitet hibridan | Visok | **Delimično** — klasa ne garantuje, ali mapper puni ID+broj iz istog reda (modSEFMapper.bas:196-197) | P3 | Bez izmene | M |
| 9 | Klasa I/II se ne šalje u UBL | Visok | **Tačno** — `ln.klasa` se ne emituje (modSEFMapper.bas:555-582 nema klasu; ni u `naziv`, 183-188) | P3 | Ako je poslovno potrebno: dopisati klasu u `opis` u mapperu (1 linija) — odluka vlasnika | S |
| 10 | Nema validation u klasi | Srednji | **Tačno** | P3 | Prihvatljivo (mapper validira) | M |
| 11 | Nema duplicate/ownership zaštite | Srednji | **Tačno** — mapper ne deduplira PrijemnicaID | P3 | Jeftin guard: `Collection.Add key:=prijemnicaID` u mapperu | S |
| 12 | Novi objekat potpuno nevalidan | Srednji | **Tačno** (VBA default vrednosti) | P3 | Bez izmene | M |
| 13 | Field-bag dizajn | Nizak | **Tačno** | Prihvaćeno | — | S |

**Bilans:** 13/13 provereno · Tačno 10 · Delimično 2 · Dizajnersko 1 · P2×2, P3×10, Prihvaćeno×1. Konkretni dobitak: line-aritmetika re-check (#2) i UBL qty/price decimale (#4) — oba u mapperu, ne u klasi.

### FM-0044 — `clsSEFResponse.cls`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | Tri nezavisna outcome Boolean-a | Kritičan | **Tačno** (cls:14-16); jedini writer (client) ih drži konzistentno; jedina „kontradikcija": status REJECTED → `Success=True`+`Rejected=True` (modSEFClient.bas:524,546-556), namerna semantika „transport ok" | P3 | Enum uz budući refactor; odmah dokumentovati semantiku `Success` | M |
| 2 | Različiti consumer prioriteti | Kritičan | **Delimično** — redosledi zaista različiti (modSEFService.bas:197→242 Success-first; :312→346 Accepted→Success→Rejected; modSEFStatusSync.bas:63→80→97 Accepted→Rejected→Success), ali za kombinacije koje client realno proizvodi ishodi se NE razilaze | P3 | Ujednačiti redosled na Accepted→Rejected→Success pri prvoj izmeni tih grana | S |
| 3 | Nema unknown outcome tipa | Kritičan | **Tačno** — `WF_SEF_UNKNOWN` postoji u state machine (modSEFValidator.bas:19) ali response ne ume da ga izrazi | **P2** | Vezati uz FM-0045 #4: unknown/blank status → `apiStatus="UNKNOWN"` → WF_SEF_UNKNOWN | S |
| 4 | Mutable posle parsiranja | Kritičan | **Tačno** kao činjenica; service tu mutabilnost i KORISTI kao guard (modSEFService.bas:172-184 prepisuje response) | P3 | Dizajnerski prihvaćeno u single-writer toku | M |
| 5 | HTTP/API/business status nepovezani | Visok | **Tačno** kao odsustvo u klasi (writer konzistentan) | P3 | Bez izmene | M |
| 6 | apiStatus nekanonski string | Visok | **Delimično** — parser normalizuje `UCase$/Trim$` (modSEFClient.bas:526; modSEFStatusSync.bas:49); fail-open je zaseban nalaz (FM-0045 #4) | P3 | — | S |
| 7 | DocumentID invariant van klase | Visok | **Tačno** — submit guard u service (modSEFService.bas:172-184); cancel/storno bez ekvivalenta | P3 | Post-cancel verify pokriva (FM-0045 #2) | S |
| 8 | Nema request/submission identiteta | Visok | **Tačno** — caller drži `submissionID` lokalno; single-thread → pogrešno uparivanje malo verovatno | P3 | Polje `requestId` na response-u pri sledećoj izmeni klase | S |
| 9 | Jedna klasa za submit/status/cancel/storno | Visok | **Tačno** | P3 | Prihvatljivo uz dokumentovanu semantiku | M |
| 10 | Success sa errorom i obrnuto | Visok | **Tačno** — konkretno postoji: status REJECTED = `Success=True` + error polja (modSEFClient.bas:550-556) | P3 | Pokriva enum refactor (#1) | M |
| 11 | Default instanca liči na failure | Visok | **Tačno** — ali client uvek popuni bar `httpStatus`/`apiStatus` (i EH putevi, modSEFClient.bas:56-63) | P3 | Bez izmene | S |
| 12 | Correlation ID nije obavezan | Srednji | **Tačno** — submit ga čak briše (modSEFClient.bas:466 → FM-0045 #10) | P3 | Vidi FM-0045 #10 | S |
| 13 | Raw body nije obavezan | Srednji | **Tačno** — EH put postavlja `""` (modSEFClient.bas:63) | P3 | Bez izmene | S |
| 14 | Nema timestamp metadata | Srednji | **Tačno** — potvrđeno: `SubmittedAt` i `FinishedAt` oba = `Now` (modSEFPersistance.bas:329-330) | P3 | Ako treba merenje: postaviti SubmittedAt pre poziva | S |
| 15 | Raw body retention/security | Srednji | **Tačno** — pun body se čuva u tblSEFSubmission; rast workbook-a | P3 | Opciono: truncate na npr. 32K pri upisu | S |
| 16 | Field-bag dizajn | Nizak | **Tačno** | Prihvaćeno | — | S |

**Bilans:** 16/16 provereno · Tačno 14 · Delimično 2 · P2×1, P3×14, Prihvaćeno×1. Nijedna „kontradikcija" nije dostižna sa postojećim writerom; realna rupa je nepostojanje UNKNOWN ishoda (#3), koja se rešava zajedno sa FM-0045 #4.

### FM-0045 — `modSEFClient.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | WinHTTP COM Windows-only | Kritičan | **Dizajnersko ograničenje** — činjenica tačna (modSEFClient.bas:314), ali CEO projekat je Windows-only: WinHttp u 9 modula (modDrive, modGoogleAuth, modMasterSync, modMonitoring, modLicense…), poppler `.exe`, WScript.Shell; „potvrđena Mac upotreba" nije proveriva u kodu i suprotna arhitekturi | Prihvaćeno | Ništa; event. capability-poruka ako se Mac ikad pojavi | L |
| 2 | Cancel/storno: svaki 2xx = uspeh, fallback CANCELLED/STORNO | Kritičan | **Tačno** — modSEFClient.bas:171-175 (cancel), :247-251 (storno); service veruje `Success` (modSEFService.bas:468,499,547,578) | **P1** | Bez fallback-a: ako body nema parsabilan `Status` → `UNKNOWN` + odmah `GetInvoiceStatus` verifikacija pre lokalnog terminalnog stanja | M |
| 3 | Submit 409 → REJECTED | Kritičan | **Tačno** — `Case 400, 409, 422` → `Rejected=True` (modSEFClient.bas:473-476); retry reuse-uje isti requestId (modSEFService.bas:37-38), pa duplicate-conflict trajno označi fakturu REJECTED iako dokument postoji na SEF → korekcioni tok → rizik duple fakture ka poreskoj | **P0** | Izdvojiti `Case 409`: `Success=False, Rejected=False, apiStatus="CONFLICT"` → service ga vodi u TECH_FAILED/manual review, ne u REJECTED | S |
| 4 | Blank/unknown status → Success/SENT | Kritičan | **Tačno** — modSEFClient.bas:597-600 (`FirstNonEmpty(statusValue,"SENT")`, `Success=True` sa :524); viši sloj unknown mapira u WF_SEF_SENT (modSEFStatusSync.bas:144-149) | **P1** | Unknown/blank → `apiStatus="UNKNOWN"` → `WF_SEF_UNKNOWN` (stanje VEĆ postoji, modSEFValidator.bas:19) + review flag | S |
| 5 | Manual JSON parser (escaped/nested/malformed) | Kritičan | **Tačno** — ExtractJsonString skenira do prvog `"` bez escape logike (modSEFClient.bas:646-649); smoke suite sam dokumentuje slabost (:975). = registrovani dug **TL-001** | Prihvaćeno (TL-001) | Ne eskalirati; pri TL-001 sanaciji escape-aware skener + testovi | M |
| 6 | Returned DocumentID prepisuje requested bez poređenja | Kritičan | **Tačno** — modSEFClient.bas:531-533 | **P2** | Ako `InvoiceId` iz body-ja ≠ requested → ne prepisivati, upisati error/UNKNOWN | S |
| 7 | Live test šalje realnu fakturu | Kritičan | **Tačno** — `Test_SubmitUBLInvoice` gradi `FAK-00001` i šalje na konfigurisani endpoint bez env guard-a (modSEFClient.bas:932-963) | **P1** | Guard: abortirati ako `SEF_ENV` nije demo/sandbox (3 linije) ili preseliti u modSEFTests | S |
| 8 | Nema typed retryability/kategorije | Visok | **Tačno** — svi EH → `HTTP_EXCEPTION` (modSEFClient.bas:52-65,111-125,189-200,265-276); service sve ne-Rejected padove → TECH_FAILED (modSEFService.bas:265) | **P2** | U EH mapirati `Err.Number` (ERR_SEF_CONFIG/ERR_SEF_VALIDATION/ostalo) u različit `errorCode` | S |
| 9 | Submit 2xx veruje HTTP statusu, ne schemi | Visok | **Delimično** — tačno (modSEFClient.bas:457-459), ali ublaženo service guard-om „success bez DocumentID = failure" (modSEFService.bas:172-184) | P2 | Zajedno sa #6 (echo/ID provera) | S |
| 10 | Correlation/remote broj se namerno gube | Visok | **Tačno** — modSEFClient.bas:465-466 postavlja `""` | P3 | Pokušati parse umesto brisanja (sadržaj submit body-ja nije statički potvrđen) | S |
| 11 | Encoding nije byte-controlled | Visok | **Delimično / nije proverivo statički** — `Send` prima String uz `charset=utf-8` header (modSEFClient.bas:33-35); WinHTTP BSTR→UTF-8 konverzija je dokumentovano ponašanje | P3 | Sandbox test sa š/ž/č u nazivu kupca; ADODB.Stream tek ako test padne | S |
| 12 | Sinhroni HTTP blokira UI | Visok | **Tačno** — `Open …, False` (modSEFClient.bas:32,92,159,235); timeouts ograničavaju blokadu (modConfig.bas:718-721: 10/10/30/30s) | P3 | Dizajnerski prihvatljivo za VBA; event. „šaljem…" status pre poziva | M |
| 13 | Cancel/storno bez operation ID/idempotency | Visok | **Tačno** — body samo invoiceId+komentari (modSEFClient.bas:154-155,229-231), URL bez requestId (:426-448) | P2 | Ako SEF API podržava requestId za cancel/storno — dodati; inače post-verify statusom (uz #2) | M |
| 14 | `CANCELED` nije poznat | Visok | **Delimično** — client zna samo `CANCELLED` (modSEFClient.bas:573), ALI unknown grana propušta statusValue u `apiStatus` (:600) a statusSync podržava OBA spelling-a (modSEFStatusSync.bas:118,257,454) → ispravno obrađeno downstream | P3 | Dodati `Case "CANCELED"` u client radi simetrije (1 linija) | S |
| 15 | Exception uvek → HTTP_EXCEPTION | Visok | **Tačno** — isti mehanizam kao #8 | P2 | Isti fix kao #8 | S |
| 16 | Debug response izlaže podatke | Visok | **Delimično** — gated `SEF_DEBUG_LOG=DA` (modSEFClient.bas:360), ide u lokalni Immediate prozor, API key se NE loguje (:384) | P3 | Prihvatljivo; event. maskirati PIB u ispisu | S |
| 17 | 429 bez strukturiranog RetryAfter | Srednji | **Tačno** — samo u poruci (modSEFClient.bas:922-928) | P3 | Uz buduću retry politiku | S |
| 18 | GUID-like validator preširok/preuzak | Srednji | **Tačno** — `1-2-3-4-5` prolazi, ne-hex slova padaju (modSEFClient.bas:883-907); napomena: greška je fail-closed (validacija odbije → nema slanja) | P3 | Zameniti stvarnim GUID regex-om + opaque fallback tek ako se pojavi realan ID | S |
| 19 | Nema content-type/schema provere | Srednji | **Tačno** — deo TL-001 paketa | P3 (TL-001) | Uz TL-001: proveriti `Content-Type` pre parsiranja | S |
| 20 | Nema TLS/host/redirect politike | Srednji | **Tačno** kao odsustvo — samo SetTimeouts (modSEFClient.bas:316-319); HTTPS enforced (:298-301); WinHTTP default već blokira https→http redirect | P3 | Bez izmene | M |
| 21 | Nema response size limita | Srednji | **Tačno** | P3 | Truncate pri persistenciji (FM-0044 #15) | S |
| 22 | Parser testovi ne pokrivaju kritične slučajeve | Srednji | **Tačno** — suite pokriva 5 osnovnih; escaped se samo ispisuje, ne assertuje (modSEFClient.bas:965-983) | P3 | Uz TL-001: assertovati escaped/blank-status/409 slučajeve | S |

**Bilans:** 22/22 provereno · Tačno 17 · Delimično 4 · Dizajnersko 1 · **P0×1 (409→REJECTED)**, P1×3 (cancel/storno false-success; blank→SENT; live test), P2×5, P3×11, Prihvaćeno×2 (Mac; TL-001). Svih 7 „Kritičan" je citirano; jedan je dizajnerski (platforma), jedan je registrovani dug TL-001.

### FM-0046 — `modSEFTax.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | Jedna hardkodovana stopa 10% | Kritičan | **Tačno** (modSEFTax.bas:3); domenski danas ispravno (otkup voća = 10%); već REGISTROVANO: selidba u tblSEFConfig | **P2** | Izvršiti registrovanu selidbu: `GetConfigValue("SEF_TAX_PERCENT")` sa fallback 10 + validacija opsega | S |
| 2 | Nema effective-date modela | Kritičan | **Tačno** kao odsustvo — funkcija ne prima datum (modSEFTax.bas:2-4); rizik se aktivira tek promenom zakonske stope; self-update flote ublažava zastarele buildove | P3 | Uz #1; datumska tabela stopa tek kad zatreba | M |
| 3 | Nema poreskog snapshot-a | Kritičan | **Delimično** — build i serialize čitaju istu konstantu u ISTOM procesu (modSEFMapper.bas:42 i :387-389); konstanta se ne može promeniti usred procesa (FM to i priznaje); poslati XML se arhivira | P3 | Snapshot stope u DTO uz FM-0042 refactor | M |
| 4 | Kategorija uvek `S` | Visok | **Dizajnersko ograničenje** (modSEFTax.bas:7) — jedan poreski režim, bez oslobođenja u domenu | P3 | Uz #1 preseliti i kategoriju u config | S |
| 5 | Stopa i kategorija nisu validirane zajedno | Visok | **Tačno** kao odsustvo — konstante su konzistentne po konstrukciji; rizik nastaje tek sa config selidbom | P3 | Pri selidbi (#1): odbiti stopa=0 uz kategorija=S bez exemption koda | S |
| 6 | Unknown business case fail-open na 10% | Visok | **Delimično** — u domenu ne postoji „druga roba" (samo otkup voća); fail-closed nema šta da hvata danas | P3 | Uz #1 validacija configa | S |
| 7 | Globalno za sve seller-e | Visok | **Dizajnersko ograničenje** — app je single-seller po workbook-u (ceo `SELLER_*` config model, modSEFMapper.bas:355-361) | P3 | Ništa | M |
| 8 | Period code `35` bez dokumentacije | Srednji | **Tačno** — nijedan komentar u fajlu (modSEFTax.bas:10-12) | P3 | Dodati ASCII komentar sa SRBDT/UBL izvorom koda 35 | S |
| 9 | Period code nije vezan za period | Srednji | **Tačno**; isti tip prometa za sve fakture → praktično konstanta | P3 | Uz #8 dokumentovati uslov važenja | S |
| 10 | Promena zahteva deployment | Srednji | **Tačno** — isto jezgro kao #1 (registrovano) | **P2** | Rešava selidba u config (#1) | S |
| 11 | Nema `Option Explicit` | Nizak | **Tačno** — fajl počinje direktno funkcijom (modSEFTax.bas:1-2); jedini SEF modul bez njega | P3 | Dodati 1 liniju uz prvi sledeći dodir fajla | S |
| 12 | `Default` naziv sugeriše nepostojeći override | Nizak | **Tačno** | Prihvaćeno | Naziv postaje tačan posle selidbe u config | S |

**Bilans:** 12/12 provereno · Tačno 8 · Delimično 2 · Dizajnersko 2 · P2×2, P3×9, Prihvaćeno×1. Tri „Kritičan" su u suštini JEDAN registrovani dug (hardkod → tblSEFConfig); dokument ga triplira umesto da referencira.

---

**Ukupno: 77/77 stavki provereno protiv koda.** Netačnih nalaza nema — FM tačno opisuje kod, ali sistematski preskače postojeće mitigacije (service DocumentID guard, arhivirani request XML pri retry-ju, dual-spelling u statusSync, WF_SEF_UNKNOWN u validatoru) i single-writer kontekst, pa su DTO „Kritični" pretežno P3. Stvarni prioriteti: **P0** — 409→REJECTED (modSEFClient.bas:473-476, fix S); **P1** — cancel/storno 2xx fallback (171-175, 247-251), blank/unknown→SENT (597-600 + modSEFStatusSync.bas:144-149), live-submit makro bez sandbox guard-a (932-963). Manual JSON parser ostaje TL-001 (ne eskalirati).

---

## Delta blok 5 — Monitoring, log, main, shell, splash (FM-0047…FM-0052, 106 stavki) [sidro f6313dc]

### FM-0047 — `modMonitoring.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | Async Send se tretira kao delivery success | Kritičan | **Tačno** — :347-352 (async, „NE pozivamo WaitForResponse"), :361 `=True` | Prihvaćeno (KR-003) | Ništa; semantika dokumentovana | — |
| 2 | Limit 12 izbacuje možda aktivne zahteve | Kritičan | **Tačno** — :355-359 pozicijsko `Remove 1` bez ReadyState provere | Prihvaćeno (KR-003) | Opciono: preskočiti objekat sa `ReadyState<4` | S |
| 3 | Nema outbox/retry — mrežni kvar trajno gubi telemetry | Kritičan | **Tačno** — jedini EH put :364-366 | Prihvaćeno (KR-003) | Outbox tek ako monitoring postane audit kanal | L |
| 4 | ActiveWorkbook config fallback | Kritičan | **Tačno — već registrovano kao AUD-018** (:455-459) | P2 | Ukloniti fallback | S |
| 5 | Windows-only WinHTTP | Kritičan | **Dizajnersko ograničenje** — cela app je Windows-only | Prihvaćeno | Ništa | — |
| 6 | Default `userId=Operator`, `role=Admin` | Visok | **Tačno** — :57-58, :40 | P3 | Uz AUTH: default iz `modAuth.GetCurrentUserIme()` | S |
| 7 | Nema event/operation ID | Visok | **Tačno** — `BuildBaseJson` :387-408 | Prihvaćeno (bez retry nema dedupe) | EventID uz eventualni outbox | S |
| 8 | Monitoring secret u svakom body-ju | Visok | **Delimično** — :389; GAS ne izlaže custom headere → platformsko ograničenje | P3 | Dugoročno HMAC(timestamp+body) | M |
| 9 | Payload sanitizacija može pokvariti JSON | Visok | **Tačno** — payload 1001–3000 znakova: :573 seče na 1000 + `[TRUNCATED]` usred JSON-a (:557), raw u envelope :407; 3000-grana jedina konzistentna | P3 | Isti replacement-objekat i za >1000 | S |
| 10 | `BuildBaseJson` fail-silent | Visok | **Dizajnersko ograničenje** — :381 OERN = „nikad ne obori business" | Prihvaćeno | Ništa | — |
| 11 | Nema client timestamp/timezone | Visok | **Tačno** — :387-408 bez `eventAt`; `IsoNow` bez TZ | P3 | `"eventAt":IsoNow()` + offset | S |
| 12 | Nema self-monitoring brojača | Visok | **Tačno** | P3 | Brojači + `Monitoring_DiagnoseConfig` ispis | S |
| 13 | Shutdown nema flush | Visok | **Tačno** — flush ne postoji | Prihvaćeno (KR-003) | — | — |
| 14 | Config lookup zaobilazi modConfig | Srednji | **Delimično** — :472-500 namerno samostalan; ipak krši reuse | P3 | Delegirati na `GetConfigValue` uz OERN | S |
| 15 | Sanitizer pokriva samo 4 tokena | Srednji | **Tačno** — :552-555 | P3 | Dopunjavati listu | S |
| 16 | Correlation generičan | Srednji | **Tačno** — :252, :176-178 | P3 | Timestamp/attempt sufiks | S |
| 17 | `sourceSpreadsheetId` = workbook name | Srednji | **Tačno** — :234 | P3 | Preimenovati polje | S |
| 18 | Backup status binaran | Srednji | **Tačno** — :223-229 | P3 | Tek uz partial semantiku | S |
| 19 | Severity/eventType nevalidirani | Srednji | **Tačno** — :392 | P3 | Allowlist normalizacija | S |
| 20 | Debug response parser substring | Srednji | **Tačno** — :327 whitespace JSON → false negative | P3 | Tolerantniji match | S |
| 21 | AppVersion fallback kružan | Nizak | **Tačno** — :29 + :432-440 mrtav kod | P3 | Obrisati fallback | S |

Bilans: 21 — 17 Tačno (1 već registrovan AUD-018), 2 Delimično, 2 Dizajnersko; 0 Netačno. 4/5 „Kritičan" u prihvaćenom KR-003 okviru. Nov S-fix: truncation 1001–3000 pravi nevalidan JSON (#9).

### FM-0048 — `modMonitoringTests.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | Async testovi „PASS bez potvrde" | Kritičan | **Delimično** — tekst je uslovna instrukcija operateru („PASS ako vidiš…"), ne lažni verdikt | P3 | Preformulisati u „PROVERI: …" | S |
| 2 | Suite može aktivirati produkcione alarme | Kritičan | **Tačno** — `TestMonitoring_All` šalje CRITICAL na realan endpoint; env se ne proverava | P2 | Env≠DEV/TEST → Yes/No potvrda | S |
| 3 | Nema burst testa (limit 12) | Visok | **Tačno** | P3 | Uz izmenu transporta | S |
| 4 | Nema shutdown/flush testa | Visok | **Tačno** | P3 | — (KR-003) | S |
| 5 | Nema sanitizer/PII testa | Visok | **Tačno** | P3 | 1 test: secret → `[REDACTED_*]` | S |
| 6 | Nema HTTP failure matrice | Visok | **Tačno** | P3 | Ručni test dovoljan za smoke | S |
| 7 | Config test ne potvrđuje source workbook | Visok | **Tačno** — root cause AUD-018 | P3 | Rešava AUD-018 | S |
| 8 | `TestMonitoring_All` bez zbirnog rezultata | Visok | **Tačno** — :4-20 | P3 | Boolean povratak | S |
| 9 | Nema backend payload assertions | Visok | **Tačno** | P3 | Van dometa VBA smoke | M |
| 10 | Test događaji zagađuju metrike | Srednji | **Delimično** — `TEST-*` ID-jevi omogućuju backend filter | P3 | Dashboard filter | S |
| 11 | Stabilni correlation ID-jevi | Srednji | **Tačno** — :63-64, :78, :82 | P3 | `Format(Now)` sufiks | S |
| 12 | Public makroi — slučajno pokretanje | Srednji | **Tačno** | P3 | Pokriveno #2 potvrdom | S |
| 13 | Nema partial backup testova | Srednji | **Tačno** — ali semantika ne postoji u modelu | P3 | — | S |
| 14 | Nema mock transporta | Srednji | **Dizajnersko ograničenje** — ručni smoke, VBA bez DI | Prihvaćeno | — | — |
| 15 | Debug.Print-only stil | Nizak | **Dizajnersko ograničenje** | Prihvaćeno | — | — |

Bilans: 15 — 11 Tačno, 2 Delimično, 2 Dizajnersko; 0 Netačno. Akcioni: env guard (#2, P2/S).

### FM-0049 — `modLogError.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | Logging failure nevidljiv | Kritičan | **Tačno** — blanket OERN :50; eksplicitni dizajn („darf NIEMALS blockieren" :14) | P3 | `m_LastWriteOk` + health check ispis | S |
| 2 | Nema sanitizacije — secret/PII u fajlu za support | Kritičan | **Tačno** — :65-84 raw; fajl je namenjen slanju supportu (:7-8) | P2 | `Monitoring_SanitizeText` → Public/shared; provući message/details | S |
| 3 | Log nije audit trail | Visok | **Dizajnersko ograničenje** — support log | Prihvaćeno | — | — |
| 4 | Newline/pipe injection | Visok | **Tačno** — :75-79; višelinijski `Err.Description` lomi format | P3 | Escape newline + `\|` | S |
| 5 | Nema correlation/build konteksta | Visok | **Tačno** | P3 | BUILD_SHA u LogAppStart (v. #11) | S |
| 6 | Concurrency dve instance | Visok | **Delimično** — single-writer; per-line Open/Close minimizuje prozor | P3 | Ništa | — |
| 7 | Nema size rotation | Visok | **Tačno** — samo 30-dnevni purge :135-171 | P3 | Opciono max-size | S |
| 8 | SOURCE sečen na 30 | Srednji | **Tačno** — :67 + PadRight | P3 | Širina 40–45 | S |
| 9 | Purge skenira sve `*.log` | Srednji | **Tačno** — :150 bez prefiksa (folder je app-ov) | P3 | `LOG_PREFIX & "*.log"` | S |
| 10 | Locale-dependent datum parse | Srednji | **Tačno** — :157-158; ISO u praksi radi | P3 | Strogi parse | S |
| 11 | App start bez BUILD_SHA | Srednji | **Tačno** — :121-125 | P3 | 1 linija `LogInfo` | S |
| 12 | Lokalni/remote log nepovezani | Srednji | **Tačno** | P3 | Uz EventID (FM-0047 #7) | M |
| 13 | `LogErr` zavisi od živog Err | Srednji | **Tačno** — :97-105; konkretan izgub u modMain (FM-0050 #12) | P3 | Fix na call-site | S |
| 14 | Prazan `ThisWorkbook.Path` | Srednji | **Tačno** — :52; .xlsm uvek snimljen | P3 | Ništa | — |
| 15 | Nema ms/timezone | Nizak | **Tačno** — :65 | P3 | — | — |
| 16 | Nemački komentari | Nizak | **Tačno** | P3 | Ne dirati | — |

Bilans: 16 — 14 Tačno, 1 Delimično, 1 Dizajnersko; 0 Netačno. P2: bez redakcije u support logu (#2, S kroz reuse).

### FM-0050 — `modMain.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | Init failure ne blokira StartApp | Kritičan | **Tačno** — InitApp EH :220-223 bez rethrow; :32 ne proverava `m_Initialized` | P2 | Posle :32: `If Not m_Initialized Then Visible=True: Exit Sub` | S |
| 2 | Missing tabele ne blokiraju init | Kritičan | **Tačno** — `ValidateAllTables` :305-308 samo MsgBox; :211-212 `m_Initialized=True` | P2 | Function Boolean; missing → ne setovati Initialized | S/M |
| 3 | Setup gate opcion | Kritičan | **Dizajnersko ograničenje** — :67-71 „ponudi, fail-soft" eksplicitno | Prihvaćeno | — | — |
| 4 | Startup EH ne vraća visibility | Kritičan | **Delimično** — rethrow :183 stiže u ThisWorkbook EH koji vraća Visible; rupa samo na splash putanji (FM-0052 #3) | P3 | Defense-in-depth `Visible=True` i ovde | S |
| 5 | Runtime schema fail-soft | Visok | **Dizajnersko ograničenje** — :193-209 namerni self-heal | P3 | — | — |
| 6 | `STARTAPP_SUCCESS` prerano | Visok | **Tačno** — :126-137 pre schedulera/shell-a | P3 | Premestiti iza | S |
| 7 | Shutdown guard bez reseta posle failure | Visok | **Tačno** — :229 True; EH :257-261 bez reseta → X postaje no-op | P2 | `mIsShuttingDown=False` u EH | S |
| 8 | `UnloadAllUserForms` beskonačna petlja | Visok | **Delimično** — rizik postoji, ali nijedna forma danas ne cancel-uje `vbFormCode` | P3 | Bounded petlja | S |
| 9 | `SaveApp` ostavlja ScreenUpdating=False | Visok | **Tačno** — :281-285 bez EH | P2 | EH/Cleanup (3 linije) | S |
| 10 | `ValidateAllTables` lista nepotpuna | Visok | **Tačno** — :290-294: 15 tabela; bez SEF/Banka/Cenovnik/Magacin; uklj. legacy TBL_CONFIG | P3 | Dopuniti iz Ensure* manifesta; izbaciti TBL_CONFIG | M |
| 11 | Init greška nelogovana | Visok | **Tačno** — EH :220-222 bez LogErr/Monitor | P2 | `LogError` pre MsgBox | S |
| 12 | StartApp EH gubi originalni Err | Visok | **Tačno — i jače nego FM**: OERN na :168 sâm briše Err → `LogErr :180` UVEK vidi 0 → lokalni log nikad ne beleži StartApp greške (snapshot :164-166 neiskorišćen). Ista klasa kao AUD-017 | P2 | :180 → `LogError "modMain.StartApp", errDesc, errNo` | S |
| 13 | App state hardkodovan pri restore | Srednji | **Delimično** — app poseduje ceo Excel proces | P3 | Ništa | — |
| 14 | Journal warning ne blokira | Srednji | **Dizajnersko ograničenje** — :106-124 warn+continue | P3 | — | — |
| 15 | Statična correlation/user | Srednji | **Tačno** — :29, :136, :175; `Operator` :24 | P3 | Session sufiks | S |
| 16 | Maintenance failure ruši startup | Srednji | **Tačno — već registrovano kao AUD-017** (:91-95) | P2 | Po AUD-017 | S |
| 17 | `gKpiDirty` globalni Boolean | Srednji | **Dizajnersko ograničenje** — :11-14 dokumentovan | Prihvaćeno | — | — |
| 18 | Nema startup state machine | Srednji | **Tačno** — :8-9 | P3 | Uz #1/#2; nice-to-have | M |
| 19 | Open/CloseExcel bez policy | Nizak | **Tačno** — :273-279; UI dugme ima auth | P3 | Ništa | — |

Bilans: 19 — 12 Tačno (1 već AUD-017), 3 Delimično, 4 Dizajnersko; 0 Netačno. Paket: fail-closed init (#1/#2) + S-fixevi (#7, #9, #11, #12 — #12 je garantovan gubitak, ne potencijalan).

### FM-0051 — `frmOtkupAPP.frm`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | Bank import se izvršava PRE auth provere | Kritičan | **Tačno** — `btnBanka_Click`: `ImportBankaInbox_WithDrivePull` :728 (Drive pull + upis) pre `OpenContentForm` :732 čiji je guard tek :1072-1077; i EH grana :754 ponovo otvara | **P1** (uz AUTH_ENABLED=YES) | Na vrh handlera blok iz `btnSyncPWA_Click` :766-772 (`AuthEnabled`→`KorisnikImaPravo`) | S |
| 2 | Initialize failure ne blokira shell | Kritičan | **Tačno** — EH :57-59 bez unload/flaga | P2 | `mSetupDone` flag; EH → `Unload Me` | S/M |
| 3 | Stale `mActiveContent` posle Show failure | Kritičan | **Tačno** — :1100 Set pre :1103 Show; EH :1120-1132 ne resetuje → Activate :69-72 trajno krije dashboard | P2 | U EH: vratiti stari pointer | S |
| 4 | KPI uključuje stornirane | Kritičan | **Tačno — već registrovano kao AUD-015** | P2 | `ExcludeStornirano` | S |
| 5 | Novi Show pre flush/unload starog | Visok | **Tačno** — :1103 pa :1106-1109 | P3 | Redosled: flush→show→unload | S |
| 6 | `ReturnToDashboard` ne unloaduje child | Visok | **Delimično** — contract: child sam sebe zatvara (10+ formi isti obrazac) | P3 | Komentar-contract; opciono defanzivni unload | S |
| 7 | Logout gubi pointer i kad unload padne | Visok | **Tačno** — `CloseActiveContent` :1013-1023 OERN + bezuslovni Nothing; poziv :1582 | P2 | `IsFormLoaded` provera; neuspeh → ne nastaviti logout | S |
| 8 | Statičan „Online" signal | Visok | **Tačno** — :1499-1510 fiksni | P3 | Neutralan tekst ili vezati za sync | S |
| 9 | Badge failure izgleda kao 0 | Visok | **Tačno** — EH :1478-1481 → 0 | P3 | -1 → „?" | S |
| 10 | Badge sabira Error/Skip/storno | Visok | **Tačno** — :1470 broji sve `<> "da"`; 4 stanja; storno neisključen | P3 | Brojati `""` + `Error`; isključiti storno | S |
| 11 | `gKpiDirty` se čisti i na neuspeh | Visok | **Tačno** — :89-92 bezuslovno | P3 | Refresh → Function Boolean | S |
| 12 | Nema readiness guarda u navigaciji | Visok | **Tačno** — root fix u modMain (FM-0050 #1/#2) | P3 | Ne duplirati u shell-u | — |
| 13 | Marža legacy izložena | Visok | **Tačno** — :379, :509, :533, :809-811; korisnik potvrdio da se ne koristi | P3 | `btnMargin.Visible=False` (runtime) | S |
| 14 | Navigacija nije role-filtrirana | Srednji | **Dizajnersko ograničenje** — click-time guard model | P3 | Opciono hide po pravima | M |
| 15 | User switch ne osvežava UI | Srednji | **Tačno** — :1586-1588; guard i dalje štiti | P3 | Refresh KPI+badge posle logina | S |
| 16 | Modeless overlay arhitektura | Srednji | **Dizajnersko ograničenje** — VBA nema MDI | Prihvaćeno | — | — |
| 17 | PWA `DoEvents` reentrancy | Srednji | **Delimično** — sidebar disabled + skriven Excel pokrivaju | P3 | Ništa | — |
| 18 | X vs Exit različita semantika | Srednji | **Tačno** — QueryClose :215-220 bez Save/Quit; chrome uklonjen pa redak | P3 | QueryClose → isti tok kao btnExit | S |
| 19 | Integrity overlay partial-build | Srednji | **Tačno** — :133-138 proverava samo 1 referencu od 4 | P3 | Setup-done flag | S |
| 20 | KPI broji fizičke redove (multi-class) | Srednji | **Delimično** — zavisi od grain-a dokumenta po klasi | P3 | `CountDistinct` po broju | S |
| 21 | UI failure bez remote monitoringa | Srednji | **Tačno** — 0 `Monitor_*` poziva | P3 | 2-3 ključna EH | S |
| 22 | Dashboard highlight btnBlocks | Nizak | **Tačno** — :96-98 | P3 | Kozmetika | S |
| 23 | Sezona = kalendarska godina | Nizak | **Tačno** — :1528 | P3 | Config po potrebi | S |

Bilans: 23 — 18 Tačno (1 već AUD-015), 3 Delimično, 2 Dizajnersko; 0 Netačno. **Najvažniji: #1 (P1) — auth-before-side-effect u btnBanka_Click**; P2 trio lifecycle (#2/#3/#7).

### FM-0052 — `frmSplash.frm`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | Shell se otvara i posle splash greške | Kritičan | **Tačno** — EH :87-89 → OpenAppShell; razuman UX izbor, readiness pripada modMain | P3 | FM-0050 #1/#2 | — |
| 2 | Shell Show retry nad istom instancom | Kritičan | **Tačno** — EH :119-123 slepi retry | P2 (paket sa #3) | v. #3 | S |
| 3 | Splash unloadovan pre potvrde shell-a → bez ijednog prozora | Kritičan | **Tačno** — :114-115 Unload pa Show; nijedna grana ne vraća `Application.Visible=True` (Excel skriven iz modMain:87); ThisWorkbook EH ne pomaže (lanac preko splash eventa) | P2 | U EH: fallback pao → `Visible=True` + MsgBox | S |
| 4 | Busy-wait 2 sekunde | Visok | **Tačno** — `WaitSeconds 2` :82, :92-106 | P3 | Smanjiti/`OnTime` | S |
| 5 | Nema readiness handoff-a | Visok | **Tačno** — :108-124 | P3 | FM-0050 fix | — |
| 6 | `m_IsNavigating` ostaje True | Visok | **Delimično** — state umire sa unloadovanom instancom | P3 | Ništa | — |
| 7 | X otvara aplikaciju | Srednji | **Dizajnersko ograničenje** — „skip wait" | P3 | — | — |
| 8 | `FindWindow` sa praznim captionom | Srednji | **Tačno** — :74 briše caption PRE :75 RemoveTitleBar | P3 | Obrnuti redosled | S |
| 9 | Fiksni 2s delay | Srednji | **Tačno/Dizajnersko** — branding | P3 | ~0.5s ili ukloniti | S |
| 10 | Splash failure bez remote monitoringa | Srednji | **Tačno** — samo LogErr | P3 | `Monitor_Error` u EH | S |
| 11 | Chrome flag bez uspeha | Nizak | **Tačno** — :76 bezuslovno | P3 | Kozmetika | — |
| 12 | Nema progress prikaza | Nizak | **Dizajnersko ograničenje** — branding splash | Prihvaćeno | — | — |

Bilans: 12 — 8 Tačno, 1 Delimično, 3 Dizajnersko; 0 Netačno. Rizik: #2+#3 (jedan S-fix) — dva pala Show-a = skriven Excel bez prozora.

**Bilans delta bloka 5 (106):** 80 Tačno (3 već registrovana: AUD-015/017/018) / 12 Delimično / 14 Dizajnersko / 0 Netačno. **Jedini P1: banka import pre auth provere** (frmOtkupAPP #1). Nova otkrića mimo FM teksta: modMain #12 — lokalni log NIKAD ne beleži StartApp greške (OERN briše Err pre LogErr); monitoring payload 1001–3000 znakova → nevalidan JSON; ValidateAllTables traži legacy TBL_CONFIG.

---

## Delta blok 6 — Auth, login, Stammdaten shell, admin/podešavanja/matični (FM-0053…FM-0063, ~166 jedinica) [sidro f6313dc]

Sve provere su završene. Evo kompletnog audita.

---

# Audit FM-0053 … FM-0063 protiv koda (worktree `f6313dc`, `src-vba/`)

Kalibracija: deployment je single-writer (jedna Excel instanca, jedan operater); auth je namerno opt-in fail-open sa anti-lockout bootstrapom (`modSetup.EnableAuth`, modSetup.bas:1426-1448, odbija bez aktivnog admina). U citatima bez naziva fajla podrazumeva se fajl te FM stavke.

### FM-0053 — `modAuth.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 55.1 | Sažetak odgovornosti | — | Kontekst — nije rizik | — | — | — |
| 55.2 | Auth fail-open po konfiguraciji (`AUTH_ENABLED` fail → sve dozvoljeno) | Kritično | **Dizajnersko ograničenje** — ponašanje potvrđeno (:34-38 EH→False; :144-147 →True), ali je dokumentovan opt-in model (:17-18) + EnableAuth bootstrap | P2 | Razdvojiti „NO" od greške čitanja: config read error → deny + poruka (tri-state po FM) | M |
| 55.3 | Unknown-form fail-open (prazna oblast → True) | Kritično | **Dizajnersko ograničenje** — potvrđeno (:207 Case Else→""; :156-159 ""→True), namera dokumentovana (:138, :193) | P2 | Nemapirana forma → deny; Debug test da je svaka navigaciona forma mapirana | S |
| 55.4 | Auth je pre svega UI-level | — | **Tačno** — jedina provera frmOtkupAPP.frm:1072-1077; javne servisne tačke (modAdmin, modPodesavanja) bez guarda | P1 | `MozeAdministraciju`/`KorisnikImaPravo` na privilegovanim javnim ulazima (vidi FM-0059/0061) | S |
| 55.5 | Duplicate username → hibridan identitet / migracija na pogrešnom redu | — | **Delimično** — login ne blokira duplikat (tačno), ali svi lookup-i su konzistentan first-match (modDataAccess.bas:439-467; `FindUserRow`=FindRows(1) :317-327) → svi čitaju ISTI prvi red, hibrid/pogrešan red ne nastaje; UI dodavanje ima dup-proveru (frmStammdaten.frm:1928-1935) | P3 | U `ValidateLogin`: `FindRows>1` → deny + log | S |
| 55.6 | Aktivnost fail-open (samo „NE" blokira) | — | **Tačno** (:86-91; komentar „drift-safe" — svesno) | P2 | Kanonizacija DA/NE pri upisu + upozorenje na nepoznatu vrednost | S |
| 55.7 | Nema brute-force zaštite | — | **Tačno** (nikakav brojač/delay u modAuth) | P2 | Failed-attempt brojač + rastući delay u `ValidateLogin` | M |
| 55.8 | Nema session timeout-a | — | **Tačno** (session do `Logout`/gašenja, :184-189) | P3 | Opcioni idle timeout | M |
| 55.9 | Plaintext fallback kad SHA ne radi | Kritično | **Dizajnersko ograničenje** — potvrđeno (:276-287; Sha256Hex EH→"" :257-258) i eksplicitno dokumentovano (:227, :232-233 „bez rizika lockout-a"); oštrica je što je TIHO | P2 | Jednokratno upozorenje + `Monitor_Event` kad `Sha256Hex` vrati prazno; `TestPinHash` pre upisa PIN-a | S |
| 55.10 | Hash slab za kratke PIN-ove (bez KDF) | — | **Tačno** (:271-287, jedan SHA-256 prolaz) | P3 | Iterirani hash ili min. dužina PIN-a (lokalni threat model) | M |
| 55.11 | Salt nije kriptografski | — | **Tačno** (`Rnd()`, :261-269) | P3 | Salt iz `Rnd`+Timer+GUID ili .NET RNG | S |
| 55.12 | Migracija fail-silent | — | **Tačno** (`MigratePinToHash` ORN, :329-334) | P3 | LogErr/Monitor na neuspeh migracije | S |
| 55.13 | Nema PIN policy-ja | — | **Tačno** (nigde min. dužina; UI traži samo neprazno) | P2 | Min. dužina u `PreparePin` + UI | S |
| 55.14 | Permission vrednosti tolerantne (DA/YES/TRUE/1/X) | — | **Tačno** (:163; sve ostalo = deny — bezbedna strana) | P3 | Kanonizacija pri upisu (UI već piše DA/NE) | S |
| 55.15 | `MozeAdministraciju` fail-open | — | **Dizajnersko ograničenje** (:173-175; komentar :170-172 — bootstrap pre uključenja; EnableAuth traži admina) | P2 | Posle prvog uključenja auth-a vezati i za trajni flag | S |
| 55.16 | Drift `OblastiList` ↔ `OblastZaFormu` | — | **Tačno** (12 oblasti :178-182 vs 10 formi :195-209; nema testa) | P3 | Provera u health check-u / `EnsureKorisniciSchema` | S |
| 55.17 | Audit best-effort, slabo strukturiran | — | **Tačno** (`AuditAuth` ORN :351-363; pri neuspehu `userId/entityID`=prazan `gCurrentUser`, pokušani user samo u msg :79,:88,:94; fiksni `VBA-AUTH` :362) | P3 | Proslediti pokušani username kao userId/entityID; per-login correlation ID | S |
| 55.18 | Logout bez audita i cleanup-a | — | **Tačno** (:184-189 samo 4 promenljive) | P3 | `AuditAuth "AUTH_LOGOUT"` u `Logout` | S |
| 55.19 | Pozitivni nalazi | — | Kontekst-Pozitivno — potvrđeno (deny bez prijave :148-151, fail-closed EH :165-167, self-test :307-315) | — | — | — |
| 55.20 | Hardening prioriteti | — | Kontekst — lista preporuka, nije nalaz | — | — | — |

**Bilans:** 20 stavki — 11 Tačno, 1 Delimično (55.5: „hibridni identitet" ne stoji zbog konzistentnog first-match), 4 Dizajnersko ograničenje (55.2, 55.3, 55.9, 55.15 — sve dokumentovan opt-in/anti-lockout dizajn), 3 Kontekst. FM „Kritično" ocene su ponašajno tačne, ali su sva tri kritična nalaza svestan dizajn → realna hitnost P2 (tihi plaintext fallback zaslužuje signal odmah — najjeftinija delta).

### FM-0054 — `frmLogin.frm`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 56.1 | Stvarna odgovornost | — | Kontekst — nije rizik | — | — | — |
| 56.2 | Generička failure poruka | — | Kontekst-Pozitivno (potvrđeno :63 — ista poruka za sva 3 tipa neuspeha) | — | — | — |
| 56.3 | „3 pokušaja" nije lockout | Kritično | **Tačno** — `mAttempts` u instanci (:19, reset :42), posle 3. samo Hide (:57-61), `modAuth.Login` unload (modAuth.bas:55) → novi ciklus. Nijansa: u startup toku 3. neuspeh vodi u `QuitAfterFailedLogin` (modMain.bas:60-62) → app se zatvara, brute-force je spor ali neograničen | P2 | Trajni brojač+delay u `modAuth` (npr. tblLocalConfig), ne u formi | M |
| 56.4 | Treći neuspeh bez posebne poruke | — | **Delimično** — forma ćuti (:57-61 Hide pre lblErr), ali odmah sledi MsgBox „prijava neuspešna" + zatvaranje (modAuth.bas:216-220 preko modMain.bas:62) — korisnik dobija generičan ishod, ne razlog | P3 | Poruka „dostignut limit" pre Hide | S |
| 56.5 | Initialize potpuno fail-silent | — | **Tačno** (:22 ORN preko cele procedure) | P3 | Izdvojiti funkcionalne korake (PasswordChar, reset) iz ORN | S |
| 56.6 | PIN može ostati nemaskiran | Kritično | **Tačno** kao mehanizam (:22 + :36 — `PasswordChar` unutar fail-silent bloka), ali scenario malo verovatan (kontrola iz .frx, prost assignment) — FM težina precenjena | P3 | `PasswordChar` prvi, van ORN; ako padne — blokirati unos | S |
| 56.7 | PIN u kontroli do success/cancel unload-a | — | **Tačno** (:49-52, :70-73 Hide bez brisanja; unload odmah u Login — kratak prozor) | P3 | `txtPin=""` pre svakog `Me.Hide` | S |
| 56.8 | Fokus na username umesto PIN | — | **Tačno** (:64) | P3 | `txtPin.SetFocus` kad je username popunjen | S |
| 56.9 | Enter/Escape nepotvrđeno | — | **Nije proverivo statički** (Default/Cancel žive u binarnom .frx; u .frm nema KeyDown handlera — tačno) | P3 | Potvrditi u dizajneru forme | S |
| 56.10 | QueryClose bez `Cancel=True` | — | Kontekst-Pozitivno — FM sam konstatuje da je funkcionalno u redu (:75-77) | — | — | — |
| 56.11 | Cancel i X ista semantika | — | Kontekst-Pozitivno | — | — | — |
| 56.12 | EH može ostaviti formu zaglavljenu | — | **Delimično** — EH samo LogErr (:66-67) tačno, ali `ValidateLogin` ima svoj EH (vraća False), kontrole ostaju aktivne → forma nije zaglavljena, samo bez feedback-a | P3 | U EH: lblErr poruka + očisti PIN | S |
| 56.13 | Nema double-click/reentrancy guarda | — | **Delimično** — činjenica, ali single-thread + sinhroni lokalni lookup → dvoklik = 2 uzastopna pokušaja (broje se), ne paralelizam; „duplicate migracija" idempotentna | P3 | Disable OK tokom validacije (higijena) | S |
| 56.14 | Nema busy/progress stanja | — | Kontekst — FM sam kaže da trenutno nije problem | — | — | — |
| 56.15 | `LoginOK` javno mutable | — | **Tačno** (:18); bez efekta na `modAuth` session (gLoggedIn ostaje False) — FM korektno ograđen | P3 | Private + read-only property | S |
| 56.16 | Nema Caps/Num Lock ili PIN pravila | — | **Tačno** (trivijalno; kozmetika) | P3 | Opciono prikaz preostalih pokušaja | S |
| 56.17 | Nema password reveal | — | Kontekst — nije rizik | — | — | — |
| 56.18 | Pozitivni nalazi | — | Kontekst-Pozitivno | — | — | — |
| 56.19 | Hardening prioriteti | — | Kontekst | — | — | — |

**Bilans:** 19 stavki — 7 Tačno, 3 Delimično, 1 Nije proverivo, 8 Kontekst/Pozitivno. Oba „Kritično" postoje u kodu, ali: 56.3 je ublažen time što startup tok zatvara aplikaciju posle 3. neuspeha, a 56.6 je malo verovatan degradacioni scenario — realna težina Srednja.

### FM-0055 — `frmStammdaten.frm`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 57.1 | Glavna uloga | — | Kontekst — nije rizik | — | — | — |
| 57.2 | `m_SetupDone=True` pre setup-a | Kritično | **Tačno** (:67-68 flag pre Setup*/LoadList; EH :159-161 ga ne resetuje → sledeći Activate izlazi na polu-formi) | P2 | Flag postaviti tek posle uspešnog setup-a; u EH `m_SetupDone=False` | S |
| 57.3 | Unknown Tag fail-open → CRUD nad Kooperantima | — | **Delimično** — fallback postoji (:134 `Case Else: SetupKooperanti`): lista tblKooperanti (PII!) se prikaže i soft-delete radi (`OnSoftDeleteClick` ne proverava Tag :1193-1241), ali Dodaj/Izmeni imaju SVOJ Case Else error (:2341-2343, :2814-2816) → nije pun CRUD | P2 | `Case Else: Err.Raise` (fatalno) | S |
| 57.4 | Nema sopstveni auth guard | — | **Tačno** — nijedna provera u formi; shell guard je propušta jer `OblastZaFormu("frmstammdaten")=""` (modAuth.bas:207 + frmOtkupAPP.frm:1072-1077) | P1 | U Activate za Tag Admin/Podešavanja/Korisnici → `MozeAdministraciju` gate | S |
| 57.5 | Add tokovi netransakcioni | — | **Tačno** (jedan pozicijski `AppendRow` :2347 bez tx za 11 mastera); ublaženo: ceo red = jedan upis | P3 | Postepeno prevesti na Korisnici obrazac (by-name + tx) | M |
| 57.6 | Korisnici dobar izuzetak | — | Kontekst-Pozitivno — potvrđeno (:1950-1983 snapshot + by-name + rollback :2368) | — | — | — |
| 57.7 | Kulture partial insert | Kritično | **Tačno** (:2184-2185 AppendRow jezgro, pa 5× običan `UpdateCell` bez tx i bez provere :2188-2193; „Dodato" :2194 i uz tihi neuspeh; na grešku EH bez rollback-a jer `korTx=Nothing`) | P2 | Isti tx obrazac kao Korisnici (`AddTableSnapshot`+`RequireUpdateCell`) | S |
| 57.8 | `GetNextID` concurrency | — | **Tačno** kao mehanizam (GetNextID→AppendRow bez claim-a, npr. :1871+:2347), ali deployment je single-writer → praktično neaktivno | P3/Prihvaćeno | Dokumentovati single-writer pretpostavku | S |
| 57.9 | Duplicate zaštita neujednačena | — | **Tačno** — username (:1928-1935, :2476-2484) i tip-šifarnici (:2216, :2238, :2260, :2282, :2297) da; Kooperant/Stanica/Kupac/Vozač/Parcela/Artikal/Kultura ne | P2 | Dup-provera po prirodnom ključu za preostale mastere | M |
| 57.10 | Soft delete TX dobar, bez provere zavisnosti | — | **Tačno** (:1221-1227; nema provere otvorenih dokumenata/salda) | P3 | Upozorenje ako master ima otvorene blokove/dug | M |
| 57.11 | Izmene uglavnom transakcione | — | Kontekst-Pozitivno (sve Edit grane `clsTransaction`+`RequireUpdateCell`) | — | — | — |
| 57.12 | Stanica lažno „Izmenjeno!" | — | **Tačno** — `UpdateFirstExistingCol` vraća Boolean, pozivi ga ignorišu (:2458-2464; definicija :1246-1256); commit + „Izmenjeno!" :2822 i kad je polje preskočeno | P2 | Za obavezna polja: `If Not UpdateFirstExistingCol(...) Then Err.Raise` | S |
| 57.13 | Istorijski drift master izmena | Kritično | **Delimično** — mogućnost izmene ovde potvrđena (sve Edit grane + tare); retroaktivni efekat zavisi od (ne)snapshot-ovanja u drugim modulima — u ovom fajlu neproverivo, FM se poziva na ranije prolaze | P2 | Efektivno-datirane tare / snapshot pri dokumentu | L |
| 57.14 | Tare se prepisuju umesto verzionisanja | — | **Tačno** (in-place Tezina: :2722-2723, :2744-2745, :2766-2767, :2788-2789) | P2 | Vidi 57.13; kratkoročno upozorenje pri izmeni težine | S |
| 57.15 | Cenovnik pravilno append-only | — | Kontekst-Pozitivno (:145 Izmeni sakriven; :2807-2812 odbija; `AddCena` :2330) | — | — | — |
| 57.16 | Nevalidan datum cene → danas | — | **Tačno** (:2326-2327); ublaženo pre-popunjenim današnjim datumom (:1143) | P3 | MsgBox na neparsiran datum umesto tihe zamene | S |
| 57.17 | Kupac validator preslab | — | **Tačno** (samo Naziv :1986-1990; PIB/MB/račun/email bez provere) | P2 | Format-provera PIB/MB pri unosu (SEF downstream) | M |
| 57.18 | PII i PIN otvoreno prikazani | — | **Tačno** (lista Kooperanti: Pin/Račun/Adresa/JMBG :1324-1333, headeri :204-213; Stanice/Vozači PIN plain txt) | P2 | Maskirati PIN kolone, JMBG delimično; role-gate prikaza | M |
| 57.19 | Korisnički PIN nasleđuje plaintext fallback | — | **Tačno** (`PreparePin` :1966, :2501 → nasleđe 55.9) | P2 | Rešava se signalom iz 55.9 | S |
| 57.20 | Poslednji admin nezaštićen | — | **Tačno** (Izmena :2486-2513 bez provere; deaktivacija poslednjeg admina uz uključen AUTH = niko ne može administraciju do ručne intervencije u sheet-u) | P2 | Blokirati deaktivaciju/demote poslednjeg aktivnog admina | S |
| 57.21 | Deaktivacija ne opoziva session | — | **Tačno** (nigde `Logout`; session je module-level) | P3 | Pri `Aktivan=NE` za `gCurrentUser` → `Logout` | S |
| 57.22 | Combo bez filtera neaktivnih | — | **Tačno** (Parcele kooperanti direktno iz tabele :845-856; kulture `GetLookupList` bez `onlyActive` :858-865 — default False, modDataAccess.bas:507-511; `LoadStaniceIntoCombo` :3454-3473) | P2 | `onlyActive:=True` / filter STATUS_NEAKTIVAN | S |
| 57.23 | Row map može postati stale | — | **Tačno** kao mehanizam (:2977-2989 fizički redovi bez PK revalidacije), single-writer + modalni tok → nisko | P3 | Pre commita uporediti PK ćeliju sa očekivanim ID-em | S |
| 57.24 | Najvažniji prioriteti | — | Kontekst | — | — | — |

**Bilans:** 24 stavke — 15 Tačno, 2 Delimično (57.3 — Add/Izmeni ipak blokirani; 57.13 — cross-module deo neproveriv ovde), 7 Kontekst/Pozitivno. Od tri „Kritično": 57.2 i 57.7 potvrđeni (jeftine S popravke), 57.13 delimično. Najveća stvarna rupa forme je 57.4 (bez auth guarda) u sprezi sa FM-0059/0061.

### FM-0056 — `clsStmBtn.cls`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 58.1 | Stvarna odgovornost | — | Kontekst — nije rizik | — | — | — |
| 58.2 | Minimalan event adapter | — | Kontekst-Pozitivno | — | — | — |
| 58.3 | Ownership problem za više instanci | Kritično | **Delimično** — poziv globalne predeclared instance je činjenica (:23), ali app koristi isključivo predeclared `frmStammdaten` (frmMaticniPodaci.frm:171; wrapper pravi ta ista instanca, frmStammdaten.frm:1176-1186) → „pogrešna forma" traži `New` instancu koje u kodu nema; FM težina precenjena | P3 | Owner referenca pri attach-u (future-proof) | S |
| 58.4 | Parent nije injektovan | — | **Tačno** (nema Owner/callback) | P3 | Uz 58.3 | S |
| 58.5 | Public mutable `btn` | — | **Tačno** (:19) | P3 | Attach metoda umesto javnog polja | S |
| 58.6 | Nema null/disposed provere | — | **Tačno** (trivijalno; bez živog dugmeta event ne postoji) | P3 | — | S |
| 58.7 | Nema error handlera | — | **Tačno**; bez posledice — `OnSoftDeleteClick` ima sopstveni EH (frmStammdaten.frm:1194, 1235-1241) | P3 | — | S |
| 58.8 | Nema reentrancy/double-click guarda | — | **Tačno** kao mehanizam; flip-flip moguć samo sekvencijalno (2 tx + 2 poruke), podaci ostaju konzistentni | P3 | Disable dugmeta tokom akcije | S |
| 58.9 | Nema auth/readiness provere | — | **Tačno** — sve zavisi od parenta koji nema service guard (57.4) | P2 | Rešiti u parent sloju (57.4), ne u wrapperu | S |
| 58.10 | Nema cleanup/Detach | — | **Tačno**; `m_softWrap` se čisti sa formom | P3 | — | S |
| 58.11 | Hardening prioriteti | — | Kontekst | — | — | — |

**Bilans:** 11 stavki — 7 Tačno (uglavnom trivijalne činjenice, P3), 1 Delimično (58.3 — jedini „Kritično", precenjen za singleton app), 3 Kontekst. Jedina stavka sa stvarnom težinom je 58.9, a rešava se u FM-0055/0059 sloju.

### FM-0057 — `modPodesavanja.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 59.1 | Stvarna odgovornost | — | Kontekst — nije rizik | — | — | — |
| 59.2 | Data-driven registar | — | Kontekst-Pozitivno | — | — | — |
| 59.3 | Partial-save model | Kritično | **Tačno** (:643-664 upis odmah po polju; int-nevalidno se preskače i javlja tek na kraju :654-655, :671-676; nema pre-validacije/tx). Pošto setteri dižu grešku (modConfig.bas:857-859; modSetup.bas:478-480), pad u sredini prekida ostatak → hibridno stanje potvrđeno | P2 | Prvo validirati SVA polja, pa tek onda pisati; opciono snapshot obe config tabele | M |
| 59.4 | Setter rezultat se ne proverava | — | **Delimično** — brojač raste bez provere (:658-662) tačno; ali oba settera re-raise na neuspeh → „prijavljen uspeh bez upisa" praktično ne nastaje (pad obara Save sa error porukom) | P3 | Pokriveno staging modelom iz 59.3 | S |
| 59.5 | Nema dirty tracking-a | — | **Tačno** (sva polja se upisuju svaki put) | P3 | Pisati samo izmenjena polja | S |
| 59.6 | Nema optimistic concurrency | — | **Tačno**; single-writer → praktično neaktivno | P3/Prihvaćeno | — | S |
| 59.7 | Secret polja obični TextBox | Kritično | **Tačno** — `"secret"` pada u istu granu kao text (:360, :376-385, bez `PasswordChar`); `SEF_API_KEY` :126, `MONITORING_SECRET` :134, `GOOGLE_CLIENT_SECRET` :142 | P2 | `If typ="secret" Then tb.PasswordChar=ChrW(8226)` — jednolinijska delta | S |
| 59.8 | Bezbednosni komentar delimično tačan | — | **Tačno** (hint :287 vs vidljivi license key/endpoints/računi :118-124 i dr.) | P3 | Ažurirati hint tekst | S |
| 59.9 | Javni `ShowConfigSheet` bypass | Kritično | **Tačno** (:725-731 Public bez ijedne provere/audita); nijansa: komentar :24 ga definiše kao namerni izlaz u nuždi — ali modSetup analogne Alt+F8 komande VEĆ guard-uje sa `MozeAdministraciju` (modSetup.bas:1362, 1429, 1453…) pa je izuzetak nekonzistentan | P2 | Isti `MozeAdministraciju` gate na `ShowConfigSheet`/`ToggleConfigSheet` (postojeći obrazac) | S |
| 59.10 | `VeryHidden` nije security granica | — | **Tačno** (konceptualno; VBE/macro vraća sheet) | P3 | Prihvatiti kao UX barijeru | — |
| 59.11 | Bool/list combo editable | — | **Tačno** (`fmStyleDropDownCombo` :362) | P3 | `fmStyleDropDownList` (pažnja na polja gde je prazno legitimno) | S |
| 59.12 | Bool formati neujednačeni | — | **Tačno** (YES/NO :364 vs DA;NE :90, :160; čitaoci tolerantni) | P3 | Jedan kanonski format pri upisu | S |
| 59.13 | Int validacija preslaba | — | **Tačno** (samo `IsNumeric` :654) | P3 | Ključ-specifični min/max u registru | S |
| 59.14 | Nema URL/PIB/račun validacije | — | **Tačno** (sve „text") | P3 | Tip „url"/„pib" u registru | M |
| 59.15 | Nema cross-field invarianta | — | **Tačno** | P3 | Minimalno: SEF_ENV↔BASE_URL upozorenje | M |
| 59.16 | Runtime efekat samo za miš | — | **Tačno** (:667-669) | P3 | „Restart potreban" napomena u poruci o čuvanju | S |
| 59.17 | Nema config audita | — | **Tačno** (nema Monitor_Event u Save) | P3 | Monitor_Event sa listom promenjenih ključeva (bez secret vrednosti) | S |
| 59.18 | Jedan Save, dve tabele, bez TX | — | **Tačno** (:657-661) | P3 | Pokriveno 59.3 | S |
| 59.19 | Legacy migracija netransakciona | — | **Tačno** (ORN :214-229; „sva 4 nova prazna" :216-219 → partial postaje trajan) | P3 | Skinuti ORN — jednokratna migracija sme da pukne glasno | S |
| 59.20 | Build može ostaviti host polukreiran | — | **Tačno** (:246-251 hide-first; EH :413-415 bez restore) | P3 | U EH pozvati `ReturnToDashboard` | S |
| 59.21 | Module-level singleton state | — | **Tačno** (:27-40); app je singleton | P3 | — | S |
| 59.22 | Poppler picker odstupa od staged Save | — | **Tačno** — `SetupPopplerInteractive` odmah persistuje (modSetup.bas:305-310, :330), folder picker samo puni polje (:485-495) | P3 | Napomena u UI ili odložiti upis do Save | S |
| 59.23 | Nema dirty prompta | — | **Tačno** (:683-694 direktan unload) | P3 | Potvrda pri Povratak ako ima izmena | S |
| 59.24 | `ApplyDefaultProizvod` fail-silent | — | **Tačno** (ORN :809-816, bez provere da vrednost postoji u combo) | P3 | Log pri neuspehu postavljanja | S |
| 59.25 | Pozitivni nalazi | — | Kontekst-Pozitivno | — | — | — |
| 59.26 | Hardening prioriteti | — | Kontekst | — | — | — |

**Bilans:** 26 stavki — 20 Tačno, 1 Delimično (59.4 — setteri ipak dižu grešku), 5 Kontekst/Pozitivno. Sva tri „Kritično" potvrđena; najbolji odnos cena/efekat: PasswordChar za secret (59.7, 1 linija) i `MozeAdministraciju` na `ShowConfigSheet` (59.9, postojeći obrazac iz modSetup).

### FM-0058 — `clsConfigBtn.cls`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | Nema editor owner/generation identiteta | Visok | **Delimično** — činjenica (cls:24-31 bez ownera), ali rebuild resetuje `mWrappers` (modPodesavanja.bas:242) → stari WithEvents sink umire i staro dugme NE MOŽE da okine; scenario suštinski neutralisan | P3 | Owner ref pri attach-u (future-proof) | S |
| 2 | Save dvoklik → dva parcijalna Save toka | Visok | **Delimično** — guarda nema (činjenica), ali Save je sinhron single-thread → dvoklik = 2 uzastopna identična upisa (idempotentno), a završni MsgBox guta drugi klik; „preklapanje" ne postoji | P3 | Disable Save tokom rada | S |
| 3 | Release/unload race | Visok | **Delimično** — `Podesavanja_Release` (modPodesavanja.bas:821-830) ubija sinkove ZAJEDNO sa referencama → klik posle release ne stiže do routera; `SaveConfigEditor` dodatno guard-uje `mInputs Is Nothing` (:635). Realna posledica: mrtva dugmad, ne rad nad praznim state-om | P3 | — | S |
| 4 | Public mutable action/groupKey | Srednji | **Tačno** (cls:24-25); eskalacija nije veća od već javnog `ConfigEditor_OnClick` | P3 | Attach metoda + private polja | S |
| 5 | Unknown action fail-silent | Srednji | **Tačno** (`ConfigEditor_OnClick` bez `Case Else`, modPodesavanja.bas:421-436) | P3 | `Case Else: LogErr` | S |
| 6 | Nema Attach/Detach contract-a | Srednji | **Tačno** (direktan assignment, `WireButton` modPodesavanja.bas:756-762) | P3 | — | S |

**Bilans:** 6 redova — 3 Tačno, 3 Delimično. Sva tri „Visok" reda su precenjena: reset kolekcije ubija stare sinkove, pa stale-click i race scenariji ne postoje u praksi; ostaje kozmetika P3.

### FM-0059 — `modAdmin.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | Nema service-level admin auth-a | Kritičan | **Tačno** — `BuildAdminPanel` (:39-52) i `AdminPanel_OnClick` (:200-221) bez ijedne provere; uzvodni guard postoji SAMO za „Korisnici" (modMaticniLookups.bas:254-259), a shell propušta frmStammdaten (`OblastZaFormu`="" → frmOtkupAPP.frm:1072-1077) → **svaki korisnik sa pravom „Matični podaci" stiže do Admin panela** | **P1** | `If Not modAuth.MozeAdministraciju() Then Exit` na vrhu obe procedure (presedan: modSetup.bas:1429) | S |
| 2 | VBA import/export/VBE u produkcionom panelu | Kritičan | **Tačno** (router :210-212; import Yes/No :296-305; export/VBE bez potvrde) | P2 | Gate iza `MozeAdministraciju` (#1) + config flag za dev komande | S |
| 3 | Fleet-wide publish sa plaintext šifrom | Kritičan | **Tačno** — `RELEASE_PUBLISH_SIFRA` je compile-time konstanta u modConfig.bas:21; InputBox nemaskiran, bez limita/audita (:277-293) | P2 | Uz #1; šifru u config/hash + audit publish-a | M |
| 4 | Cleanup/migracija bez maintenance lock-a | Kritičan | **Delimično** — direktno routovanje tačno (:213-214) i lock-a nema, ali VBA single-thread znači da tokom sinhrone komande ista instanca ne može paralelno raditi poslovne operacije; multi-instance nije podržan model (single-writer) | P3 | Osloniti se na postojeće interne potvrde; opciono „operacija u toku" flag | S |
| 5 | Ensure agregat može lažno uspeti | Visok | **Tačno** kao mehanizam (:257-272 — uspeh se izvodi iz odsustva exceptiona); fail-soft prirode pojedinačnih `Ensure*` ovde nisu re-verifikovane (FM se poziva na raniji audit) | P3 | `Ensure*` da vraćaju status; agregat da sumira | M |
| 6 | Partial runtime panel build | Visok | **Tačno** (:47-52 hide-first; EH :145-147 bez rollback-a; `m_SetupDone` već True — frmStammdaten.frm:68) | P3 | U EH `ReturnToDashboard` | S |
| 7 | Nema structured result/audit-a | Visok | **Tačno** (sve `Sub` bez rezultata) | P3 | Postepeno; prvo za destruktivne komande | M |
| 8 | Nema double-click/reentrancy guarda | Visok | **Delimično** — činjenica (:119-125 bez disable), ali sinhrono izvršavanje + modalne potvrde (publish/import/migracija/cleanup) ograničavaju realan rizik | P3 | Disable panela tokom akcije | S |
| 9 | `OnTime` scheduling se ne potvrđuje | Visok | **Tačno** (:243; EH :251-252 samo log — panel već zatvoren, korisnik bez poruke) | P3 | MsgBox u EH `AdminCheckUpdate` | S |
| 10 | Offline prikazan kao „nema update-a" | Srednji | **Delimično** — poruka RAZLIKUJE stanja u tekstu (:246-248 napomena „kanal nije dostupan"), ali headline „koristite najnoviju verziju" je offline neproveren | P3 | Odvojena poruka za `remote=""` | S |
| 11 | Unknown action fail-silent | Srednji | **Tačno** (:202-216 bez `Case Else`) | P3 | `Case Else: LogErr` | S |
| 12 | Singleton owner state | Srednji | **Tačno** (:33-34) | P3 | — | S |
| 13 | Close failure gubi controller reference | Srednji | **Tačno** (:308-314 ORN — ako `Unload mFrm` padne, reference se ipak brišu) | P3 | Proveriti unload pre brisanja referenci | S |
| 14 | Samo lokalni error log | Srednji | **Tačno** (svuda `LogErr`, nigde Monitor_Event) | P3 | Monitor_Event za publish/import/cleanup neuspehe | S |

**Bilans:** 14 redova — 10 Tačno, 3 Delimično, 1 Tačno-uslovno (#5). Red #1 je **najjači pojedinačni nalaz celog opsega** (potvrđen kompletan lanac do Admin panela za ne-admin korisnika) i jedini pravi P1; #4 je za single-writer deployment precenjen. Minimalna delta za #1+#2+#3: jedan `MozeAdministraciju` guard.

### FM-0060 — `clsAdminBtn.cls`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | Nema admin-session/owner validacije | Kritičan | **Tačno** kao činjenica (cls:24-30 — ništa se ne čuva ni proverava); rizik postoji isključivo zato što je router neguardovan (modAdmin.bas:200) — wrapper sam ne dodaje površinu | P1 (preko FM-0059 #1) | Guard u routeru, ne u klasi | S |
| 2 | Public mutable `action` → preusmerenje na destruktivnu komandu | Kritičan | **Delimično** — činjenica (cls:24), ali preusmerenje traži VBA kod koji ionako može DIREKTNO pozvati `OcistiTabele`; nema dodatne eskalacije → težina precenjena | P3 | Private polje + Attach (higijena) | S |
| 3 | Dvoklik/reentrancy | Visok | **Delimično** — guarda nema (činjenica), ali sve rizične akcije imaju modalne potvrde i sinhrone su (single-thread) | P3 | Disable dugmadi tokom akcije | S |
| 4 | Release/unload race | Visok | **Delimično** — `CloseAdminPanel` briše `mWrappers` (modAdmin.bas:311-313) → sink umire, klik posle release ne stiže; router akcije su ionako globalne procedure nezavisne od `mFrm` | P3 | — | S |
| 5 | Globalni singleton router | Srednji | **Tačno** | P3 | — | S |
| 6 | Nema Attach/Detach | Srednji | **Tačno** (`WireButton` modAdmin.bas:319-325 direktan assignment) | P3 | — | S |
| 7 | Nema lokalni EH context | Nizak | **Tačno**; router EH loguje action (:219) — dovoljno | P3/Prihvaćeno | — | S |

**Bilans:** 7 redova — 4 Tačno, 3 Delimično. Oba „Kritičan" su izvedenice neguardovanog routera (FM-0059 #1); sama klasa je korektan minimalni adapter, popravka pripada `modAdmin`.

### FM-0061 — `modMaticniLookups.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | Admin/Podešavanja bez guarda u meniju | Kritičan | **Tačno** — `MaticniMenu_OnClick` proverava SAMO „Korisnici" (:254-259); Admin/Podešavanja prolaze, downstream bez zaštite (modAdmin.bas:200; modPodesavanja.bas:421; frmStammdaten bez guarda; shell propušta preko `OblastZaFormu`="") | **P1** | Proširiti postojeći guard: `If sekTag="Korisnici" Or "Admin" Or "Pode<š>avanja" Then MozeAdministraciju` — 1 uslov | S |
| 2 | Registry/frmStammdaten drift → Kooperanti fallback | Kritičan | **Tačno** — dva izvora istine (komentar :14-15, :65); nemapiran Tag → frmStammdaten.frm:134 fallback (uz nijansu 57.3: Dodaj/Izmeni blokirani, lista+soft-delete rade) | P2 | Ukloniti `Case Else` fallback + Debug provera registry↔Case | S |
| 3 | Globalna predeclared owner forma | Visok | **Tačno** kao činjenica (`AttachMaticniMenu` ne čuva frm :92-97; OnClick → globalna forma :262); posledica hipotetička u singleton app | P3 | Owner ref (future-proof) | S |
| 4 | Partial build bez cleanup-a | Visok | **Tačno** (EH :210-213 ostavlja već kreirane kontrole; rebuild uklanja samo imena koja ponovo kreira :173-176) | P3 | U EH ukloniti `btnMD_*`/`lblMDgrp_*` | S |
| 5 | Statični fallback samo 6 sekcija | Visok | **Tačno** (`STATIC_BTNS` :28-29) | P3 | Prihvatiti kao degradaciju ili prikazati poruku | S |
| 6 | Singleton module state | Visok | **Tačno** (:24-26); singleton app | P3 | — | S |
| 7 | Nema permission metadata u registru | Visok | **Tačno** (:67-88 samo Caption+Tag) | P2 | Treći element `requiredAdmin` po sekciji (podloga za #1) | S/M |
| 8 | `MozeAdministraciju` nasleđuje fail-open | Srednji | **Dizajnersko ograničenje** (modAuth.bas:173-175; opt-in model + EnableAuth bootstrap) | P2 (uz 55.2) | — | S |
| 9 | Highlight pre otvaranja | Srednji | **Tačno** (:261-262; OpenSekcija EH vraća formu bez reset stila — frmMaticniPodaci.frm:186-192) | P3 | Highlight posle uspeha ili reset u EH | S |
| 10 | Nema duplicate/invalid Tag validacije | Srednji | **Tačno** (bez validacije; duplikat Tag → Remove+Add pregazi prvo dugme :173-178) | P3 | Debug provera unique Tag | S |
| 11 | Forma raste bez screen bounds | Srednji | **Tačno** (:197-204) | P3 | Clamp na Application height | S |
| 12 | Release ostavlja mrtva dugmad | Srednji | **Tačno** (:286-290 samo kolekcije; kontrole i `mHoverNm` ostaju); kontekst: release je namenski pre self-update importa | P3/Prihvaćeno | — | S |
| 13 | Hover stale posle ResetAll | Nizak | **Tačno** (`ResetAll` :228-235 ne dira `mHoverNm`; OnHover early-exit :242) — kozmetika | P3 | `mHoverNm=""` u ResetAll | S |
| 14 | Empty registry edge case | Nizak | **Tačno** (:53 `ReDim 0 To -1` bi pukao), trenutno neaktivan | P3 | Guard `If out.count=0` | S |

**Bilans:** 14 redova — 12 Tačno, 1 Dizajnersko ograničenje, 1 Tačno-latentno. Oba „Kritičan" potvrđena; #1 zajedno sa FM-0059 #1 čini glavni P1 lanac ovog opsega — popravka je bukvalno proširenje već postojećeg guarda za „Korisnici".

### FM-0062 — `clsLookupMenuBtn.cls`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | Nema owner/generation identiteta | Visok | **Tačno** kao činjenica (:25-33); posledica hipotetička — singleton meni, rebuild ubija stare sinkove (modMaticniLookups.bas:95-97) | P3 | Owner ref (future-proof) | S |
| 2 | Javno mutable `sekcijaTag` | Visok | **Delimično** — činjenica (:25), ali preusmerenje ka Admin traži VBA kod koji može i direktno `OpenSekcija "Admin"`; stvarni rizik je neguardovan router (FM-0061 #1), ne mutabilnost | P3 | Private + Attach | S |
| 3 | Tag nije validiran prema registru | Visok | **Tačno** — prosleđuje se svaki string (:31-33), router bez validacije, frmStammdaten fallback (frmStammdaten.frm:134) | P2 | Rešava se FM-0061 #2/#7 | S |
| 4 | Nema auth/session metadata | Visok | **Tačno** (permission samo u routeru i samo za Korisnici — modMaticniLookups.bas:254-259) | P2 (preko FM-0061 #1) | Guard u routeru | S |
| 5 | Caption/Tag mogu biti kontradiktorni | Srednji | **Tačno** (oba javna, bez uparivanja; caption ide u naslov — frmMaticniPodaci.frm:180) | P3 | Attach sa parom iz registra | S |
| 6 | Dvoklik/reentrancy | Srednji | **Delimično** — guarda u klasi nema, ali `OpenSekcija` radi `Unload Me` (drugi klik pada u prazno) i postoji `m_IsOpeningChild` | P3 | — | S |
| 7 | Hover nad stale globalnom kolekcijom | Srednji | **Delimično** — rebuild resetuje `mWrappers` → stari hover sink umire pre nego što može da dira novu kolekciju | P3 | — | S |
| 8 | Nema Attach/Detach | Srednji | **Tačno** | P3 | — | S |
| 9 | Nema lokalnog error contexta | Nizak | **Tačno** (click router loguje generično :266; hover ORN) | P3 | Tag u log poruci routera | S |

**Bilans:** 9 redova — 6 Tačno, 3 Delimično. Suštinski rizici (#3, #4) su isti authorization-drift problem kao FM-0061 i tamo se rešavaju; ostatak je P3 higijena singleton adaptera.

### FM-0063 — `frmMaticniPodaci.frm`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | Ključni open pozivi pod `Resume Next` | Kritičan | **Tačno** (:178-181 ORN oko `frmOtkupAPP.Show` + `OpenContentFormPublic`) | P2 | Ukloniti ORN — EH (:186-192) već ume da vrati meni | S |
| 2 | Meni se unloaduje bez potvrde child-a | Kritičan | **Tačno** (:183 bezuslovni `Unload Me`; `OpenContentFormPublic` ne vraća rezultat — guard odbijanje u frmOtkupAPP.frm:1072-1077 je tihi `Exit Sub`). Nijansa: shell ostaje vidljiv ako `Show` uspe, pa korisnik nije bez ikakvog UI-ja, ali sekcija tiho izostane | P2 | `OpenContentFormPublic` → Function Boolean; unload samo na True | M |
| 3 | Nema sopstveni permission guard | Kritičan | **Tačno** (`OpenSekcija` :155-193 bez ijedne provere; jedini uzvodni gate je OBL_MATICNI na ulasku u meni — frmOtkupAPP.frm:910) | **P1** (deo lanca FM-0059/0061) | Guard u `MaticniMenu_OnClick` (FM-0061 #1) pokriva i ovo | S |
| 4 | Force unload prethodnog child-a pre validacije | Kritičan | **Tačno** (:166-168 pre svega osim blank-checka; bez dirty provere) — ali gubitak je nesnimljen unos u poljima pri svesnoj navigaciji korisnika → težina realno Srednja | P3 | Unload starog child-a tek posle uspešnog open-a novog | M |
| 5 | Unknown Tag nije validiran | Visok | **Tačno** (:158-161 samo blank check) | P2 | Validacija prema registru (uz FM-0061 #2) | S |
| 6 | Close path guta shell Show failure | Visok | **Tačno** (:116-120 ORN Show pa Unload) | P3 | Proveriti `frmOtkupAPP.Visible` pre unload-a | S |
| 7 | Partial Initialize bez readiness state-a | Visok | **Tačno** (:71-74 EH bez flag-a/unload-a; klikovi rade dalje) | P3 | `mSetupFailed` flag + blok klikova | S |
| 8 | Globalne predeclared forme | Visok | **Tačno** (:167, :171, :179-180); singleton dizajn cele aplikacije | P3 | — | S |
| 9 | Nema structured open result/rollback | Visok | **Tačno** (`OpenSekcija` je Sub) | P3 | Uz #2 | M |
| 10 | Deactivate zatvara meni na focus loss | Srednji | **Dizajnersko ograničenje** — namerno popup ponašanje (:90-101 sa flagovima) | Prihvaćeno | — | — |
| 11 | Highlight pre uspeha | Srednji | **Tačno** (:241-269 `ButtonActive` pre open-a; EH re-Show bez reseta) | P3 | Reset stila u EH | S |
| 12 | Static fallback nepotpun | Srednji | **Tačno** (6 dugmadi :59-64) | P3 | Vidi FM-0061 #5 | S |
| 13 | Caption-based HWND sa praznim captionom | Srednji | **Tačno** (:81-82 caption="" pa `RemoveTitleBar`; `FindWindow("ThunderDFrame", "")` :34; i frmStammdaten briše caption :153 → dva prazna captiona moguća) | P3 | Privremeni jedinstveni caption pre FindWindow | S |
| 14 | Pogrešno ime u log source-u | Srednji | **Tačno** (`frmStammdatenMenu.*` :72, :87, :100, :124, :147 vs `frmMaticniPodaci.OpenSekcija` :187 — mešano) | P3 | Ujednačiti na stvarno ime forme | S |
| 15 | Nema remote monitoring-a | Srednji | **Tačno** (samo LogErr) | P3 | Monitor_Event za navigation failure | S |
| 16 | Ponovljen chrome API na Activate | Nizak | **Tačno** — namerno (komentar :79-80 „ne koristiti mChromeRemoved"); kozmetički trošak | Prihvaćeno | — | — |

**Bilans:** 16 redova — 14 Tačno, 1 Dizajnersko ograničenje, 1 Prihvaćeno (namerni pattern). Sva 4 „Kritičan" postoje u kodu; #3 je deo glavnog P1 lanca, #1/#2 su realni P2 (tihi gubitak sekcije), #4 je precenjen (Srednji).

---

## Ukupni zaključak

Od ~166 stavki: ogromna većina je činjenično tačna prema kodu; oko 20 je **Delimično** (najčešće zato što reset wrapper-kolekcija ubija stale event sinkove, što FM nije prepoznao, ili zbog single-writer/single-thread realnosti), a fail-open auth stavke su **dokumentovan dizajn** (opt-in + `EnableAuth` anti-lockout), ne propust. **Jedini pravi P0/P1 nalaz je lanac autorizacije:** korisnik sa pravom „Matični podaci" stiže do Admin panela (Očisti tabele, Migracija, VBA Import/Export, Publish sa šifrom `"agrix-release"` iz modConfig.bas:21) jer modMaticniLookups.bas:254-259 guard-uje samo „Korisnici", modAdmin.bas:39/200 i modPodesavanja.bas:725 nemaju sopstvenu proveru, a shell guard propušta `frmStammdaten` (modAuth.bas:207 → frmOtkupAPP.frm:1072-1077). Minimalna delta: proširiti postojeći `MozeAdministraciju` guard na „Admin"/„Podešavanja" u `MaticniMenu_OnClick` + isti jednolinijski gate na vrhu `BuildAdminPanel`/`AdminPanel_OnClick`/`BuildConfigEditor`/`ShowConfigSheet` (obrazac već postoji u modSetup.bas:1429). Sledeći po prioritetu (P2, sve S/M napor): PasswordChar za „secret" polja, Kulture insert u transakciju, `m_SetupDone` posle uspešnog setup-a, signal pri plaintext PIN fallbacku, uklanjanje `Case Else→Kooperanti` i ORN-a oko child open-a.

---

## Delta blok 7 — Pregled listova, migracija, journal, lock, orchestrator, sync, sheets, auth, drive (FM-0064…FM-0072, 126 jedinica) [sidro f6313dc]

All verification is complete. Here is the full audit report.

# Audit FM-0064 — FM-0072 protiv koda (worktree `wt-f6313dc/src-vba/`, commit `f6313dc`)

Svi navodi `fajl:linija` odnose se na `/tmp/claude-0/-home-user-otkupapp-pwa/c27e5940-dcae-584b-9571-644cbf8a2f95/scratchpad/wt-f6313dc/src-vba/`. Kalibracija: single-writer desktop; multi-writer važi samo na cloud/PWA sync površinama. Pre-registrovani nalazi: AUD-001 (P0, JSON read parser), AUD-006 (journal), AUD-019 (PIN/JMBG export), KI-006 (ART_POCETNI_DUG), prior P2 (SyncControl RMW/last-writer-wins).

### FM-0064 — `modPregledListova.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 66.1 | Stvarna odgovornost: servisni indeks + launcheri | — | **Tačno** (kod odgovara opisu) | — | — | — |
| 66.2 | Public API: 5 procedura | — | **Tačno** (23, 81, 92, 104, 117) | — | — | — |
| 66.3 | Nema autorizacije ni za jednu akciju | Kritično | **Tačno** kao činjenica (nigde provere usera), ali Alt+F8 svakako izvršava svaki makro — VBA guard je samo savetodavan | P3 | List držati VeryHidden + advisory provera role pre `OcistiTabele` | S |
| 66.4 | Lista sve sheetove uklj. hidden/VeryHidden | — | **Tačno** (42-54, bez `Visible` filtera) | P3 | Preskočiti `xlSheetVeryHidden` | S |
| 66.5 | Persistentna privilegovana površina | — | **Tačno**, ali sheet nastaje samo ručnim `NapraviPregledListova` (komentar 11) | P3 | Isto kao 66.3 | S |
| 66.6 | `PokreniProgram` zaobilazi startup gate | — | **Delimično** — identičan obrazac kao ostali pozivaoci (`frmMarza`…); StartApp je već prošao pri otvaranju (komentar 78-80); gap samo ako je first-run setup odbijen | P3 | Provera `APP_SETUP_COMPLETED` pre `Show` | S |
| 66.7 | `OtvoriVBA` javna dev komanda | — | **Dizajnersko ograničenje** — Alt+F11 je nativno uvek dostupan; dugme ne dodaje privilegiju (92-100) | P3 | Ukloniti dugme ili flag | S |
| 66.8 | `SendKeys` fallback nepouzdan | — | **Tačno** (97) | P3 | Poruka umesto SendKeys | S |
| 66.9 | `PokreniMigraciju` bez guarda | — | **Tačno** (106); potvrda u modMigracija postoji samo ako tblOtkup ima redove | P3 | Suština u FM-0065; ovde ništa | — |
| 66.10 | `OcistiTabele` bez transakcije/backupa | Kritično | **Tačno** (modPregledListova.bas:117-158 — nema clsTransaction/snapshot) | P2 | `SaveCopyAs` backup pre brisanja (reuse `BackupFileOnStart` obrazac) | S |
| 66.11 | False-success helper | Kritično | **Tačno** (modPregledListova.bas:192-201 — `True` čim tabela postoji; `Delete` pod `Resume Next`) | P2 | Posle delete proveriti `DataBodyRange Is Nothing` | S |
| 66.12 | „22/22" može lažno prijaviti uspeh | — | **Tačno** (139-151, zavisi od 66.11) | P2 | Isti fix kao 66.11 | S |
| 66.13 | Missing vs failed delete nerazdvojeni | — | **Tačno** (False samo za nepostojeću tabelu) | P2 | Tri ishoda: missing/failed/cleared | S |
| 66.14 | Typed potvrda smanjuje slučajan klik | Pozitivno | **Kontekst-Pozitivno** (131-134, 162-168; prima i dijakritiku) | — | — | — |
| 66.15 | Nema backup restore point-a | — | **Tačno** — isti nalaz kao 66.10 | P2 | Spojeno sa 66.10 | S |
| 66.16 | Nema provere aktivnih operacija | — | **Tačno** (nema provere formi/sync-a) | P3 | Provera `VBA.UserForms.Count` | S |
| 66.17 | Cleanup lista nepotpuna | — | **Tačno** — 22 od ~40 `TBL_*` (modConfig.bas:37-89,312,729); ne briše tblParcele, tblStornoVeze, tblSEFSubmission/EventLog, tblBankaImport, tblKorisnici… → orphan reference | P3 | Dokumentovati obuhvat ili dopuniti listu | S |
| 66.18 | Config reference ostaju | — | **Tačno** — `DefaultVrsta/DefaultSorta/TipAmbalaze` u tblSEFConfig ostaju posle brisanja mastera | P3 | Upozorenje u završnoj poruci | S |
| 66.19 | Nema audit događaja | — | **Tačno**; journal pokriva samo AppendRow (AUD-006 kontekst) | P3 | LogInfo + Monitor_Event | S |
| 66.20 | `Cells.Clear` briše sadržaj lista | — | **Tačno** (31), ali sopstveni list, regeneracija dokumentovana | Prihvaćeno | — | — |
| 66.21 | Briše sva Forms dugmad | — | **Tačno** (204-210), samo na sopstvenom listu | P3 | — | — |
| 66.22 | `OnAction` nije workbook-qualified | — | **Tačno** (230) | P3 | `"'" & ThisWorkbook.Name & "'!Proc"` | S |
| 66.23 | Pozitivni nalazi (9) | Pozitivno | **Tačno** — svih 9 potvrđeno | — | — | — |
| 66.24 | Hardening prioriteti (20) | Hardening | **Delimično** — jezgro (backup, provera delete-a, razdvajanje ishoda, qualified OnAction) opravdano; operation ID/structured result/maintenance servis/regression testovi prekomerni za ručni dev alat | P2/P3 | Samo stavke 5, 9-11, 19 | S |

**Bilans:** 24 stavke — 19 Tačno, 2 Delimično, 1 Dizajnersko ograničenje, 2 Kontekst-Pozitivno/Prihvaćeno. Realan delta: jedan mali patch (backup + tačan rezultat brisanja + qualified OnAction); „autorizacija" je u Excelu inherentno savetodavna.

### FM-0065 — `modMigracija.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 67.1 | Best-effort transfer, ne atomic engine; hibridno stanje moguće | Kritično | **Tačno** (modMigracija.bas:66-91 — per-tabela `Resume Next`, greška u summary, bez rollback-a) | P3 | Jednokratni ručni alat sa file-pickerom; backup pre run-a dovoljan | S |
| 67.2 | 20 kritičnih nalaza | Kritično | **Tačno** u celini — svih 20 potvrđeno: upozorenje samo tblOtkup (21-30); zamena ne merge (122); clear pre mape (122 vs 125-135); nemapirana kolona preskočena (157); `StaroImeKolone` prazan (274-276); first-match (251-260, 200); secrets/tblLocalConfig se prenose (171-176, 263-266 — svesno, komentar); mapirane kalkulisane kolone prepisane vrednostima (122+162); summary bez overall statusa (106-109); ensure best-effort (46-49) | P2 (samo #2) | Minimal delta: upozorenje ako IJEDNA ne-config tabela u cilju ima redove, ne samo tblOtkup; opciono `SaveCopyAs` cilja pre run-a | S |
| 67.3 | Pozitivni nalazi (9) | Pozitivno | **Tačno** — svi potvrđeni (ReadOnly 57, makroi off 53, state restore 100-103, by-name 124-135…); poklapa se s prior auditom („by-name mapping is careful") | — | — | — |
| 67.4 | Hardening prioriteti (16) | Hardening | **Delimično** — opravdano: backup cilja (#2), šire upozorenje (#: proveriti sve tabele), politika za tblLocalConfig (#12 — sporno jer je komentar 263-266 svestan, ali per-machine putanje sa starog PC-ja jesu rizik); manifest/ledger/dry-run/atomic temp-commit prekomerno za jednokratni alat | P3 | Stavke 2, 7 (delimično), 12 | S/M |

**Bilans:** 4 stavke — 3 Tačno, 1 Delimično. Jedan S-fix vredan pažnje (upozorenje gleda samo tblOtkup pre destruktivnog prepisa).

### FM-0066 — `modJournaling.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 68.1 | CSV journal nije transaction journal; red pre commit-a | Kritično | **Tačno** — `WriteJournalRow` zvan iz `AppendRow` (modDataAccess.bas:210), bez commit/rollback markera; familija **AUD-006** | AUD-006 | Pokriveno AUD-006 remedijacijom | — |
| 68.2 | 20 kritičnih nalaza | Kritično | **Tačno** u celini; #2-#4 (samo append / danas-vs-ukupno count) = **već registrovano kao AUD-006**; ostalo potvrđeno: samo današnji fajlovi (181), multiline count (198-201), header −1 bezuslovno (203), tihi write fail (55), purge `*.csv` u Journal folderu (125), max-age tek posle prvog save-a (498), `m_SaveScheduled=True` i kad OnTime padne (531-534), `FlushNow` bez rezultata (515-521), backup minutska rezolucija (272), **backup rethrow blokira startup** (modJournaling.bas:331 `Err.Raise` + modMain.bas:91 poziv pod aktivnim `On Error GoTo EH`); #20 Delimično (IsDate na ISO string radi na većini locale-a) | P2 (#19); ostalo P3/AUD-006 | #19: policy odluka — ili `Resume Next` oko `BackupFileOnStart` u StartApp, ili zadržati blokadu ali sa jasnom porukom operateru | S |
| 68.3 | Pozitivni nalazi (12) | Pozitivno | **Tačno** — svi potvrđeni (escaping 382-390, SaveCopyAs 286, reentrancy guard 426-427, DisplayAlerts restore 450-476, `ThisWorkbook.Saved` 511/519…) | — | — | — |
| 68.4 | Hardening prioriteti (18) | Hardening | **Delimično** — #1-#4, #7 = srž AUD-006; #12-#14 (max-age pre prvog save-a, provera OnTime, FlushNow rezultat) mali i opravdani; per-instance lock/encryption/remote monitoring prekomerno za single-writer | AUD-006 + P3 | Stavke 12-14, 17 | S |

**Bilans:** 4 stavke — 3 Tačno, 1 Delimično; jezgro je već registrovano kao AUD-006, novi izdvojiv nalaz je backup-fail-blokira-startup (P2, policy).

### FM-0067 — `modStanicaLock.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 69.1 | RMW celog SyncControl taba, bez CAS/lease | Kritično | **Tačno** (modStanicaLock.bas:558-600 — read dict → full `WriteSheetData`) = **prior review finding, P2 već registrovan** | P2 (reg.) | Pokriveno postojećim P2 (GAS-side CAS / key-level update) | M |
| 69.2 | 19 kritičnih nalaza | Kritično | **Tačno** u celini: acquire ne čita postojeći lock (85-99), OWNER=`vba` konst. (38, 93), release bez owner provere (167-177), stale cleanup bez CAS (284-348), **prazan UPDATED_AT nikad ne ističe** (318 `If Len>0`), timer flag i kad OnTime padne (237-241), release ignoriše rezultat (177), append→marker nije atomsko (434-440), bez storno filtera u bulk push-u (410-447), ChangeStanica pušta stari pre potvrde novog (81-99); #18 (desktop-only True) = **Dizajnersko ograničenje** — bez clouda nema druge strane (67-72, dokumentovano) | P2 (reg.) + P3 (#16) | Uz postojeći P2: LeaseID kolona (S); #16: preskočiti stornirane redove u push petlji | S |
| 69.3 | Pozitivni nalazi (8) | Pozitivno | **Tačno** — svi potvrđeni (heartbeat 90s < TTL 10min, push pre unlock-a 162-165, marker prazan za retry 441-443…) | — | — | — |
| 69.4 | Hardening prioriteti (13) | Hardening | **Delimično** — CAS/LeaseID/atomic key update opravdani (GAS postoji, izvodljivo); TTL config/monitoring/permission guard prekomerni | P2 (reg.) | Stavke 1-4, 10 | M |

**Bilans:** 4 stavke — 3 Tačno, 1 Delimično; sve u okviru već registrovanog prior P2, novo samo storno-filter (P3/S).

### FM-0068 — `modGoogleSyncOrchestrator.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 70.1 | Public API | — | **Tačno** (46-52, 624-706, 765-775) | — | — | — |
| 70.2 | Dobar redosled + guard, ali master lock prepisuje ceo SyncControl | Kritično | **Tačno i verifikovano do kraja**: `SetPWAMasterSyncLock` piše tačno 5 redova (modGoogleSyncOrchestrator.bas:578, 590-605) kroz `WriteSheetData` koji je **full-tab staging replace** (modGoogleSheets.bas:263-357, 834-932) → **briše sve `STANICA_LOCK_*` ključeve** koje modStanicaLock čuva RMW-om (asimetrija!) | P2 | Minimal delta: SetPWAMasterSyncLock da koristi isti RMW obrazac kao `ApplySyncControlUpdates` (modStanicaLock.bas:558-600) — čuva ostale ključeve | S |
| 70.3 | 40 kritičnih nalaza | Kritično | **Tačno** u celini; ključne potvrde: outbound nastavlja nezavisno (290-299) vs inbound fail-fast; ok=samo `Err.Number` (197-204, 221-228, 250-257); Boolean AND 7 flagova (301-308); multi-poziv Monitora sa različitim corrId (456 + 12 call-site-ova); user „Operator" (480, 522); DoEvents (437); callback ni module-qualified (655, 713, 723 — za razliku od modStanicaLock/modJournaling); bez backoff-a (674-688); implicitno kreiranje Stammdaten (554); AddSheetTab progutano (560-562); #24 **Delimično** (parse fail se loguje, „tiho" samo prema korisniku, 742-754); #29 **Delimično** — `SetConfigValue` je `Sub` (modConfig.bas:796), nema rezultata koji bi se proverio | P2 (#1-3) ostalo P3 | #1-3 = predlog iz 70.2; #22: kvalifikovati callback; #26-27: brojač uzastopnih grešaka + stop posle N | S |
| 70.4 | Pozitivni nalazi (14) | Pozitivno | **Tačno** — svi potvrđeni (geo hard gate 26+159-176, unlock fail → cycle fail + poseban monitoring 345-375, min interval 15 min 29+635…) | — | — | — |
| 70.5 | Hardening prioriteti (19) | Hardening | **Delimično** — #1 (key-level update) je pravi S-fix; #3 (stabilan SyncRunID) S; #16 opravdan; distributed lease/ledger/manifest/reconciliation prekomerno za ovaj obim | P2 (#1) | Stavke 1, 3, 15, 16 | S |

**Bilans:** 5 stavki — 3 Tačno, 2 Delimično (u okviru lista). Glavni novi nalaz: master lock briše station lock ključeve — proširenje registrovanog SyncControl P2, fix je S (reuse postojećeg RMW helpera).

### FM-0069 — `modStammdatenSync.bas` (tabela 71.52, 19 redova)

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | Plaintext PIN export → cloud Users tab | Kritičan | **Tačno** (modStammdatenSync.bas:2212-2216, 2241-2245, 2274-2278, 2306-2310, 2348-2352) = **već registrovano kao AUD-019** | AUD-019 | Pokriveno AUD-019 | — |
| 2 | PII export (JMBG, adresa, telefon, BPG) | Kritičan | **Tačno** (modStammdatenSync.bas:1326-1327, 1347-1348, 1367, 1380-1381) = **AUD-019** | AUD-019 | Pokriveno AUD-019 | — |
| 3 | Neatomic 13-tab publish | Kritičan | **Delimično** — sekvenca jeste per-tab (200-212), ali per-tab staging+verify (FM-0070) štiti svaki tab; ostaje samo prolazna mešovina verzija koju sledeći ciklus ispravlja; False + CRITICAL monitoring postoji (225-231) | P3 | `ExportedAt/SyncRunID` red u Config tabu | S |
| 4 | Empty source → header-only overwrite | Kritičan | **Tačno** (modStammdatenSync.bas:1324-1337 i isti obrazac u svim exporterima; `GetTableData=Empty` ne razlikuje kvar od praznog) | P2 | Guard: ako lokal 0 redova a tab je ranije imao podatke (ili prosto min-count za Kooperanti/Users/Fakture) → abort tog taba | S/M |
| 5 | Direktni Parcele export zaobilazi geo pull gate | Kritičan | **Tačno** (modStammdatenSync.bas:122-124→202 i 1569-1616 nemaju pull; gate samo u modGoogleSyncOrchestrator.bas:26, 159-176) — javni makro može pregaziti novije PWA poligone | P2 | U `SyncStammdatenToGoogle_Core` i `SyncParceleToGoogle_Core` pozvati `ImportParcelGeoFromGoogleToMaster` pre `ExportParcele` (isti gate) | S |
| 6 | FakturaStavke bez parent/storno filtera | Visok | **Tačno** (modStammdatenSync.bas:2488-2494 — nema `ExcludeStornirano` ni provere parenta) | P3 | Filtrirati stavke po aktivnim FakturaID iz već izgrađenog skupa | S |
| 7 | Username collision | Visok | **Tačno** (modStammdatenSync.bas:2241, 2274, 2306 — `LCase(Left(ime,1)&prezime)`, bez dedup provere; PWA login dvosmislen) | P2 (deo AUD-019 remedijacije) | Kratkoročno: sufiks EntityID pri koliziji | S |
| 8 | OtkupiAll synthetic timestampi | Visok | **Tačno** (modStammdatenSync.bas:777-806 — `UpdatedAtServer=Now` 781, `ReceivedAt=Now` 798, CreatedAt prazni 779-780) | P3 | Reporting tab; dokumentovati semantiku ili preneti stvarni datum | S |
| 9 | First-match receipt join po BrojZbirne | Visok | **Delimično / Dizajnersko** — komentar 869-870 pokazuje svesnu odluku („prva je dovoljna"); rizik postoji kod duplikata BrojZbirne | P3 | Join po `PrijemnicaID` gde postoji | M |
| 10 | Nema dataset version/manifest | Visok | **Tačno** | P3 | Spojeno sa #3 predlogom | S |
| 11 | Nema row-count sanity check | Visok | **Tačno** | P3 | Spojeno sa #4 | S |
| 12 | Nema snapshot consistency | Visok | **Tačno** (svaki `GetTableData` u drugom trenutku; DoEvents u orchestratoru dozvoljava izmene) — single-writer + master lock znatno ublažava | P3 | — | — |
| 13 | Kartice nisu opening-balance saldo | Visok | **Tačno** (modStammdatenSync.bas:344-345 — od 1.1. tekuće godine; `ReportKarticaKooperanta` modIzvestaj.bas:314 računa samo period) | P3 | Poslovna odluka vlasnika: potvrditi da li PWA kartica sme biti godišnji movement | — |
| 14 | SaldoOM/Detail/Kupci ≠ centralni reporti | Visok | **Tačno** (komentar „vereinfacht" 995-996; SaldoOM samo 2 tipa novca 1029-1033; Detail 1191 bez ambalaže) | P3 | Centralizovati definiciju (jedan izvor za report i export) | M |
| 15 | `IsPWAActive` fail-open | Srednji | **Delimično** — prazno=aktivan verovatno namerno za legacy redove (2601-2611); typo rizik realan | P3 | Log upozorenja za nepoznate vrednosti | S |
| 16 | Kulture bez active/storno filtera | Srednji | **Delimično** — tblKulture po šemi NEMA `Aktivan` ni storno kolonu (CLAUDE.md), pa predloženi filteri ne postoje; „ne filtrira" tehnički tačno ali bespredmetno (1401-1425) | P3 | Eventualno unique Vrsta+Sorta provera | S |
| 17 | Locale string brojevi/datumi | Srednji | **Tačno** (`CStr` svuda, npr. 609-612, 2463; decimalni separator zavisi od locale-a) | P3 | Invariant-format helper (`Str$`/Replace) | S/M |
| 18 | Hardcoded user + statični correlation ID | Srednji | **Tačno** (2621, 2626, 2641 — `Operator`, `STAMMDATEN-SYNC`) | P3 | Per-run ID (timestamp) | S |
| 19 | Public test macro radi produkcioni overwrite | Srednji | **Tačno** (2663-2665), ali isto važi za javni `SyncStammdatenToGoogle` — namerni ručni entry | P3 | Ukloniti `Test_SyncStammdaten` (redundantan) | S |

**Bilans:** 19 redova — 14 Tačno, 4 Delimično, 1 Dizajnersko; 2 Kritična su već AUD-019, novi P2: empty-source guard (#4/#11) i geo-gate za samostalni Parcele export (#5), plus username kolizija u okviru AUD-019.

### FM-0070 — `modGoogleSheets.bas` (72.1-72.46)

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 72.1 | Glavni zaključak: hardenovan, ali swap neatomski, parser krhak | — | **Tačno** — precizna ocena | — | — | — |
| 72.2 | Public API | — | **Tačno** (uklj. javni `ParseValuesJson` 1734) | — | — | — |
| 72.3 | Staging write štiti target | Pozitivno | **Kontekst-Pozitivno** — poznata jačina (staging-verify-swap, 301-343) | — | — | — |
| 72.4 | Replace nije atomska operacija | Kritično | **Tačno** (modGoogleSheets.bas:834-932 — tri odvojena batchUpdate poziva: rename 887, rename 897, delete 917; između prva dva target ime ne postoji za PWA čitača) | P2 | Spojiti oba rename-a u JEDAN `batchUpdate` sa dva requesta (Google batchUpdate je atomski za listu) — prozor nestaje | S |
| 72.5 | Recovery posle 2. rename failure best-effort | — | **Tačno** (897-909; fail-recovery ostavlja `__old_*`) | P3 | Pada zajedno sa 72.4 fixom | — |
| 72.6 | Staging/backup ime sekundska rezolucija | — | **Tačno** (419, 429), ali kolizija traži 2 paralelna write-a iste sekunde uz throttle 1250ms — malo verovatno | P3 | Dodati `Timer`-ms sufiks | S |
| 72.7 | Nema WriteOperationID/lease | — | **Tačno**; multi-instance edge | P3 | — | — |
| 72.8 | Target sheetId se menja | — | **Dizajnersko ograničenje** — svesno dokumentovano (40-42) | Prihvaćeno | — | — |
| 72.9 | Post-replace verify bez rollback-a | — | **Tačno** (339-343), ali sadržaj je već verifikovan u staging-u; False → sledeći ciklus prepiše | P3 | — | — |
| 72.10 | Verify `>=` umesto exact | — | **Tačno** (473-485); na svežem staging tabu bezbedno | P3 | — | — |
| 72.11 | Sve string + RAW | — | **Dizajnersko ograničenje** — svesno (komentar 1695-1698: čuva vodeće nule) | P3 | — | — |
| 72.12 | Datumi gube vreme | — | **Tačno** (1718-1719 `vbDate`→`yyyy-mm-dd`); pogađa npr. `OtkupiAll` `UpdatedAtServer=Now` (modStammdatenSync.bas:781) | P3 | Ako Date ima time deo → `yyyy-mm-dd hh:nn:ss` | S |
| 72.13 | Custom JSON parser krhak | Kritično | **Tačno — već registrovano kao AUD-001 (P0)** (1734-1802) | AUD-001 | Pokriveno AUD-001 | — |
| 72.14 | Escaped quote menja quote-state | Kritično | **Tačno — već registrovano kao AUD-001** (1816) | AUD-001 | — | — |
| 72.15 | Literal `],[` razbija redove | Kritično | **Tačno — već registrovano kao AUD-001** (1781) | AUD-001 | — | — |
| 72.16 | Strip newline/space menja vrednost | Kritično | **Tačno — već registrovano kao AUD-001** (1746-1759 — uklj. `", "`→`","` i unutar navodnika) | AUD-001 | — | — |
| 72.17 | colCount iz prvog reda (ragged) | — | **Tačno** (1787-1799 truncate/pad) — isti parser, rešava se AUD-001 remedijacijom | AUD-001 | — | — |
| 72.18 | Unescape nepotpun (`\uXXXX`…) | — | **Tačno** (1832-1841; verify normalizacija 524-534 dodatno maskira korupciju) — deo AUD-001 | AUD-001 | — | — |
| 72.19 | Retry helper vraća True posle iscrpljenih pokušaja | — | **Tačno** (99-101, 114), ali svi calleri potom proveravaju `http.status` — nema false-success; problem imenovanja | P3 | Preimenovati / komentar | S |
| 72.20 | 65s DoEvents busy-wait | — | **Tačno** (145-155, 199-201) | P3 | `Sleep` API + ređi DoEvents | S |
| 72.21 | Nema cancel/deadline | — | **Tačno** | P3 | — | — |
| 72.22 | Throttle per-instance | — | **Tačno** (55, 157-179); single-writer ublažava | P3 | — | — |
| 72.23 | CreateSpreadsheet/Move/Search bez retry helpera | — | **Tačno** (1323, 1402, 1616, 1647 — direktan `Send`) | P3 | Provući kroz `SendGoogleHttpWithRetry` | S |
| 72.24 | Move fail ipak vraća ID | — | **Delimično** — tačno (1343-1353), ali ID se odmah upisuje u config pa se dalje koristi po ID-u; duplikat nastaje samo ako se config izgubi | P3 | — | — |
| 72.25 | GetSpreadsheetID first-match | — | **Tačno** (1414, 1440-1470) | P3 | Upozorenje ako query vrati >1 exact-name | S |
| 72.26 | Drive search bez paginacije | — | **Tačno** (1395-1396, `pageSize=10`, bez pageToken) | P3 | — | — |
| 72.27 | Drive metadata parser krhak | — | **Tačno** (1472-1491), ali imena fajlova su kontrolisana (`Stammdaten`…) | P3 | — | — |
| 72.28 | AddSheetTab True bez potvrđenog sheetId | — | **Tačno** (1560-1566 — rezultat force refresha se ignoriše) | P3 | `AddSheetTab = (existingSheetId > 0)` u toj grani | S |
| 72.29 | `checkExisting=False` staging bez collision zaštite | — | **Tačno** (303 + 1560-1566), ista niska verovatnoća kao 72.6 | P3 | Spojeno sa 72.6 | S |
| 72.30 | Append bez idempotency key | — | **Tačno** (1056-1145; retry na 5xx posle uspešnog server-side prijema duplira red); da li GAS/uvoz dedupira po `VBA:<OtkupID>` — **Nije proverivo statički** ovde | P3 | Proveriti GAS dedupe; ako ga nema → podići na P2 | S |
| 72.31 | Append potvrda samo HTTP 2xx | — | **Tačno** (1131-1138) | P3 | Pročitati `updates.updatedRows` | S |
| 72.32 | Append ne proverava 1D shape | — | **Delimično** — tačno da proverava samo `IsArray` (1088), ali 2D ulaz završi u runtime error → EH → False (nema false-success) | P3 | — | — |
| 72.33 | ReadSheetData `Empty` višeznačno | — | **Tačno** (1176-1227 — svi failure putevi → Empty); konkretan lanac: `ReadSyncControlAsDict` na transient fail dobije prazan dict → RMW upiše samo update ključeve → **briše ostale SyncControl ključeve** (modStanicaLock.bas:532-555) | P2 (u okviru registrovanog SyncControl P2) | Razdvojiti EMPTY od ERROR (ByRef ok flag) i u RMW abortovati na ERROR | S/M |
| 72.34 | ClearSheet javna destruktivna | — | **Tačno** (1233-1285), bez pozivalaca u hardened toku | P3 | — | — |
| 72.35 | ClearSheet nekorišćen u WriteSheetData | Pozitivno | **Kontekst-Pozitivno** — potvrđeno | — | — | — |
| 72.36 | Cache bez TTL/invalidation | — | **Tačno** (542-586; miss → force refresh 610-619 ublažava) | P3 | — | — |
| 72.37 | Cache nije distributed | — | **Tačno**; single-writer | P3 | — | — |
| 72.38 | Backup cleanup nije implementiran | — | **Tačno** (917-922 warn; 1027-1038 helper samo loguje) — `__old_*` sa PII ostaje samo kad delete padne | P3 | Allowlist cleanup `__old_/__stage_` starijih od N dana | S |
| 72.39 | Staging/backup dupliraju sensitive | — | **Tačno** (privremeno; trajno samo na delete fail) | P3 | Spojeno sa 72.38 | S |
| 72.40 | Log body može nositi sensitive | — | **Tačno** (222-224, prvih 1000 znakova; error body tipično bez values) | P3 | — | — |
| 72.41 | Nema structured result | — | **Tačno** | P3 | — | — |
| 72.42 | Nema per-request monitoring | — | **Tačno** | P3 | — | — |
| 72.43 | Nema payload/chunking policy | — | **Tačno** (1692-1732 konkatenacija — O(n²) na velikim tabovima) | P3 | `Mid$`-buffer builder ako OtkupiAll poraste | M |
| 72.44 | Full verify skup (2×GET po tabu) | — | **Tačno** — svestan trošak za integritet | Prihvaćeno | — | — |
| 72.45 | Pozitivni nalazi (20) | Pozitivno | **Tačno** — svi potvrđeni | — | — | — |
| 72.46 | Hardening prioriteti (26) | Hardening | **Delimično** — #3 (atomski dvostruki rename = 72.4/S), #9 (EMPTY vs ERROR = 72.33), #8 (parser = AUD-001), #17, #20 opravdani; lease/CAS/GoogleApiResult/monitoring stack prekomerno | P2 (#3, #9) | Stavke 3, 8, 9, 17, 20 | S/M |

**Bilans:** 46 stavki — 36 Tačno (od toga 6 = AUD-001 već registrovan: 72.13-72.18), 3 Delimično, 2 Dizajnersko/Prihvaćeno, 3 Kontekst-Pozitivno, 1 Nije proverivo statički (GAS dedupe u 72.30), 1 Prihvaćeno (72.44). Nova P2 delta: atomski rename-par (S) i EMPTY≠ERROR za ReadSheetData (S/M, vezano za SyncControl P2).

### FM-0071 — `modGoogleAuth.bas` (tabela 73.37, 16 redova)

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | OOB flow bez redirect/PKCE | Kritičan | **Tačno** (modGoogleAuth.bas:28 `urn:ietf:wg:oauth:2.0:oob`; 54-60 bez `code_challenge`/`state`) — Google je OOB zvanično ugasio, radi li još za ovaj client **nije proverivo statički**; rizik prestanka rada realan | P2 | Plan migracije na loopback redirect (+PKCE usput) | M |
| 2 | Refresh token/secret plaintext u tblSEFConfig | Kritičan | **Tačno** (43-44, 97-98, 186-190, 255-262), ali to je dokumentovan dizajn (CLAUDE.md: „kredencijali žive u tblSEFConfig"); workbook je trust boundary — token ipak daje cloud pristup širi od fajla | P2 | DPAPI/Credential Manager za refresh token (ostalo može ostati) | M |
| 3 | Full Drive scope | Kritičan | **Tačno** (29 — `auth/drive`, ne `drive.file`); migracija na `drive.file` zahteva proveru svih postojećih fajlova/foldera | P3 | Proceniti `drive.file` izvodljivost (self-update folder!) | M |
| 4 | Token save nije atomic | Kritičan | **Tačno** kao činjenica (186-190, 255-262 — više `Call SetConfigValue`), ali failure mod je lokalna tabela (retko) i oporavlja se refresh-om/re-auth-om; **napomena:** `SetConfigValue` je `Sub` (modConfig.bas:796) — nema rezultata za proveru | P3 | Upis u fiksnom redosledu: expiry poslednji; opciono setter → Function | S |
| 5 | Nema refresh concurrency lock-a | Visok | **Delimično** — VBA je single-threaded (nema „dva paralelna poziva" u instanci); svaka mašina ima svoju kopiju config tabele; scenario zahteva dve instance nad istim fajlom (druga je read-only) | P3 | — | — |
| 6 | Public setup bez admin guarda | Visok | **Tačno** (35), ali to je dokumentovani setup entry (CLAUDE.md/modAdmin) i Alt+F8 je svejedno otvoren — **Dizajnersko ograničenje** | P3 | — | — |
| 7 | Nema REAUTH_REQUIRED stanja | Visok | **Tačno** (235-240 — `invalid_grant` samo loguje; `IsGoogleAuthConfigured` ostaje True 124-129) | P3 | Na `invalid_grant` upisati flag + jasna poruka operateru | S |
| 8 | Nema 401 forced-refresh | Visok | **Tačno** (samo lokalni expiry 107; modGoogleSheets ne retry-uje na 401) | P3 | U wrapperu: 401 → jedan force refresh → retry | S |
| 9 | Nema account/scope verifikacije | Visok | **Tačno** (nigde se ne čuva email/scope) | P3 | Sačuvati `scope` iz token response + email (tokeninfo) | S |
| 10 | Token endpoint bez retry/backoff | Visok | **Tačno** (156-158, 229-231 — jedan `Send`; kontrast: modGoogleSheets ima retry helper) | P3 | Reuse `SendGoogleHttpWithRetry` obrazac | S |
| 11 | Expiry bez TZ + locale parser | Srednji | **Tačno** (346-348 lokalni `Now` bez `Z`; 336 `CDate` locale; parse fail = fail-safe expired 342-343); DST unazad → do 1h mrtvih 401 poziva (pojačava #8) | P3 | UTC + ručni ISO parse; ili samo #8 fix | S |
| 12 | Shared global credential | Srednji | **Tačno** — dizajn (jedan servisni Google nalog) | Prihvaćeno | — | — |
| 13 | Simple public JSON parser | Srednji | **Tačno** (350-390, javni, koriste ga modDrive/modGoogleSheets); tokeni su base64url pa je rizik nizak | P3 | Uz AUD-001 remedijaciju | — |
| 14 | Nema revoke/logout API | Srednji | **Tačno** | P3 | `RevokeGoogleAuth` (revoke endpoint + brisanje ključeva) | S |
| 15 | Redakcija nije strukturalna | Srednji | **Delimično** — mehanizam tačno opisan (295-311 zamena samo sačuvanih vrednosti), ali body se loguje samo na ne-200 (162-165, 235-238) gde Google ne vraća tokene → scenario „nesačuvan token u logu" praktično nedostižan | P3 | — | — |
| 16 | Nema centralnog auth audita/health | Srednji | **Tačno** (nema Monitor_* poziva u modulu) | P3 | Monitor_Event na setup/refresh fail | S |

**Bilans:** 16 redova — 12 Tačno, 2 Delimično, 1 Dizajnersko ograničenje (u okviru reda 6), 1 Prihvaćeno; od 4 „Kritična" po kalibraciji ostaju 2 P2 (OOB migracija; DPAPI za refresh token), ostalo P3 uz male S-fixeve (REAUTH flag, 401-retry, token-endpoint retry).

### FM-0072 — `modDrive.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 74.1 | Glavni zaključak: download→direktan overwrite; find ""→create duplikat | Kritično | **Tačno** — oba obrasca potvrđena (modDrive.bas:40-44 `SaveToFile destPath,2` bez temp/hash; 70-78 find vraća `""` i za HTTP error i za not-found; 91-94 `""` → `DriveCreateEmpty`) | P2 | Vidi 74.2 | S |
| 74.2 | 25 kritičnih nalaza | Kritično | **Tačno** u celini — svi potvrđeni: pageSize=1 (66), list bez paginacije + prvi ID po imenu (132, 149), split po `}` (144), bez mimeType/md5/size (62, 66, 132), bez retry/401 (svi `Send`), metadata-create pre čitanja lokalnog fajla (93 pre 97 — fail ostavlja prazan remote), self-test pravi `agrix_selftest.txt` bez brisanja (186-195), bez guarda/validacije putanje; #3 **Delimično** (sinhroni WinHttp vraća kompletan body — truncation malo verovatna; nepotvrđena je samo *ispravnost sadržaja*); #19 zero-byte svesno obrađen (243-245). **Najjači lanac (P2):** transient non-200 na find (70-74) → duplikat release fajla (91-94) → `DriveListFolder` prvi ID po imenu (149) → self-update fleet-a može povlačiti stari artefakt | P2 (lanac 4-5-14); ostalo P3 | Minimal delta u `DriveFindInFolder`: na `status<>200` vratiti error-signal (npr. `Err.Raise` ili ByRef ok=False) umesto `""`, i u `DriveUploadFile` abortovati; opciono download u `.part` + rename | S |
| 74.3 | Pozitivni nalazi (11) | Pozitivno | **Tačno** — svi potvrđeni (binarno bez transkodiranja, supportsAllDrives, trashed=false, timeouti 235…) | — | — | — |
| 74.4 | Hardening prioriteti (20) | Hardening | **Delimično** — jezgro opravdano (#6 find-fail≠not-found, #1-#3 temp+verify+rename, #14 paginacija — trenutno ~100 fajlova pa je 1000 dovoljno); appProperties ID/idempotency op ID/audit/monitoring prekomerni za obim alata | P2 (#6), P3 ostalo | Stavke 6, 1-3, 16 (self-test cleanup) | S |

**Bilans:** 4 stavke — 3 Tačno, 1 Delimično. Jedan izdvojen P2 (find-error tretiran kao not-found → duplikat release artefakta + first-match self-update), fix je S.

---

## Zbirni pregled

- **Ukupno verifikovano:** 126 jedinica (24+4+4+4+5+19+46+16+4). Ništa nije Netačno; FM je činjenično vrlo precizan — glavna korekcija je **kalibracija težine** (mnogo „Kritično" pada na P3 u single-writer kontekstu) i mapiranje na već registrovane nalaze (AUD-001 ×6 u FM-0070, AUD-006 u FM-0066, AUD-019 ×3 u FM-0069, prior SyncControl P2 u FM-0067/0068/0070).
- **Nove P2 stavke koje audit predlaže za registar (sve S/M napora):** (1) `SetPWAMasterSyncLock` full-tab overwrite briše `STANICA_LOCK_*` — fix: reuse RMW helpera (FM-0068); (2) empty-source → header-only cloud wipe guard (FM-0069 #4); (3) samostalni Parcele/Stammdaten export bez geo pull gate-a (FM-0069 #5); (4) neatomski rename-par u swap-u → jedan batchUpdate (FM-0070 72.4); (5) `ReadSheetData` EMPTY≠ERROR (72.33, u SyncControl P2 familiji); (6) `modDrive` find-error→duplikat release lanac (FM-0072); (7) `OcistiTabele` backup + tačan rezultat brisanja (FM-0064); (8) OAuth: OOB migracija + DPAPI za refresh token (FM-0071); (9) `BackupFileOnStart` rethrow policy (FM-0066 #19); (10) modMigracija upozorenje gleda samo tblOtkup (FM-0065).

---

## Delta blok 8 — Self-update, gate, licenca, trial, license testovi (FM-0073…FM-0078, 151 podsekcija) [sidra f6313dc + a0bc9e2]

Verifikacija je kompletna (svih 6 fajlova + cross-file: `modMain`, `ThisWorkbook.doccls`, `modVbaTools`, `modPodesavanja`, `modConfig`, `modPregledListova`, `gas/Code.gs`, runbook docs, git istorija). Ključne materijalne provere: `modLicense.bas` i `modTrial.bas` su **identični** u oba checkout-a (line-ref važe za oba); `ThisWorkbook.doccls` **nema** `AccessWasDenied()` proveru iako je runbook (docs/production-runbook-licenca.md:86) i komentar (modLicense.bas:75–77) tvrde; duplikat test modula postoji i razlikuje se **samo u headeru**; git klon je **shallow** (8 graft root-ova) pa se raniji „deletion commit" ne može potvrditi.

---

### FM-0073 — `modSelfUpdate.bas` (anchor f6313dc)

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 75.1 | Live updater bez atomic replace/manifest/compile-gate/auto-rollback; najopasniji tok DeleteLines→AddFromString | Zaključak (kritično) | **Tačno** — modSelfUpdate.bas:313–315 (Delete pa Add pod `Resume Next`); retry (297–335) ne vraća stari body; nema compile/manifest gate-a. Deo pre-registrovan (files_count fix planiran) | P1 (snapshot deo); manifest deo registrovan | In-memory snapshot starog body-ja po modulu + restore pri padu; manifest `files_count` (već planiran) | M |
| 75.2 | 36 kritičnih nalaza (backup ručni, parcijalni import, fiksni temp/registry ključevi, state briše se pre importa, nema mutex/potpisa/health-a…) | Kritični nalazi | **Tačno** (34/36): rollback samo instrukcija (:72–75, :124–126); Delete pre potvrde (:313–315); nastavlja od n≥1 (:79–84); temp fiksni (:252); phase-2 ključevi bez RunID (:106–107); state obrisan pre importa (:139); faza 2 samo `.bas/.cls` (:149) — **plus gore od FM-a: failed `.frm` se u fazi 1 Remove-uje (:101–105) a faza 2 ga ne uvozi → komponenta nestaje**; events ostaju off (:194–195); Save gutanje (:113–115). **Delimično** nr.30 (min-version gate postoji zasebno u modUpdateGate); nr.20 nije u celosti proverivo statički | P1: nr.2/3/11/14; P2: manifest klasa (4/5/31/32 — registrovano); P3 ostalo | Minimalni delta: snapshot+restore, durable phase-2 state (brisati TEK posle uspeha), faza 2 da pokrije i failed `.frm` (ili da ih ne Remove-uje), `EnableEvents/ScreenUpdating` restore u EH | M |
| 75.3 | 15 pozitivnih | Pozitivno | **Kontekst-Pozitivno** — svih 15 potvrđeno u kodu | — | — | — |
| 75.4 | 23 hardening prioriteta | Prioriteti | **Tačno** (konzistentno s nalazima); pun opseg je redesign | P2 (podskup) | Sprovesti samo: manifest+count (planiran), snapshot/restore, durable state; ostalo backlog | M/L |

**Bilans:** činjenično gotovo sve tačno (1 delimično, 1 neproverivo); značajan deo je registrovani dizajn dvofaznog live-update-a. Nove akcione tačke: snapshot/restore starog koda, durable phase-2 state i **rupa gde failed `.frm` biva uklonjen bez ponovnog importa** (FM je čak potcenio nr.14).

---

### FM-0074 — `modUpdateGate.bas` (anchor f6313dc)

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 76.1 | Fail-open na sve greške; `VERSION_ENFORCE=YES` nije security granica; permisivni parser verzija | Zaključak | **Tačno** činjenično + **Dizajnersko ograničenje** (fail-open 2,5 s pre-registrovan: :16–17, :31, :39–49, :79–82); parser: 3 segmenta (:137), suffix od `-`/`+` (:158–159), `Val/CLng`→0 (:163–166) | Prihvaćeno / P3 | Dokumentovati da prerelease/4. segment nisu podržani u release šemi (git-describe odsecanje je namerno, :129–131) | S |
| 76.2 | 23 kritična nalaza (bez cache-a/potpisa/HTTPS provere, unknown enforce→WARN, prazan min gasi, 2,5 s bez retry-a, isti ishod za sve greške, parser 14–19, fiksni temp, bez audita) | Kritični nalazi | **Tačno** (23/23): fail-open modUpdateGate.bas:39,43,46,49,79–82; enforce `Select Case`:61–76; `latest` samo za poruku :53–57; jedini pozivalac modMain.bas:44; timeout 2500 :31,:105; non-2xx→False→propust :110–114; manifest :184–189. Nalazi 1–3, 11–12, 18 = registrovani availability dizajn | Prihvaćeno (1–3, 11–12); P3 parser (14–19, bez aktivnog nosioca — nema prerelease tagova); P3 ostalo | Ako gate ikad postane security granica: keširana poslednja policy + vremenski ograničen fail-open; do tada samo `enforce` strict parse (`Case Else` uz WARN log) | S |
| 76.3 | 12 pozitivnih | Pozitivno | **Kontekst-Pozitivno** — svih 12 potvrđeno | — | — | — |
| 76.4 | 20 hardening prioriteta | Prioriteti | **Tačno** (konzistentno); većina menja dogovoreni availability-first model | Prihvaćeno / P3 | Eventualno samo typed rezultat + razlikovanje timeout/auth/malformed u logu | S/M |

**Bilans:** sve tvrdnje potvrđene u kodu; ništa novo van registrovanog fail-open dizajna. Parser-permisivnost je realna ali trenutno bez nosioca (release šema je `vba-vX.Y.Z`, git-describe odsecanje je dokumentovano namerno).

---

### FM-0075 — `modLicense.bas` (anchor f6313dc; fajl identičan i na a0bc9e2)

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 77.1 | Stvarni model (3 komponente, server bind, lokalni cache, trial bez ključa) | Opis | **Tačno** (:8–14, :47–54, :474–510, :84–91) | — | — | — |
| 77.2 | Licenca odvojena od cloud-sync flaga | Pozitivno | **Kontekst-Pozitivno** (:148–153) | — | — | — |
| 77.3 | Latch sprečava `LICENSE_ENABLED=NO` bypass | Pozitivno | **Kontekst-Pozitivno** (:556–567) | — | — | — |
| 77.4 | **Kritično:** `AccessGateOrQuit` exception → fail-open | Kritično | **Dizajnersko ograničenje** — modLicense.bas:126–129 (`EH: LogErr; AccessGateOrQuit = True`), komentar :127 dokumentuje odluku; unutar plafona | Prihvaćeno | — (restatement dizajna) | — |
| 77.5 | `LicenseGateOrQuit` fail-open na exception | Nalaz | **Dizajnersko ograničenje** (:277–281, komentar) | Prihvaćeno | — | — |
| 77.6 | Prazan endpoint pušta licenciranu instalaciju | Nalaz | **Tačno** + Dizajnersko (:154–161, LogWarn pa True; komentar eksplicitan) | P3 | Jednokratni MsgBox operateru umesto samo loga | S |
| 77.7 | Slab otisak → fail-open | Nalaz | **Dizajnersko ograničenje** (:171–177); sabotaža WMI je unutar plafona (VBE korisnik može više) | Prihvaćeno | — | — |
| 77.8 | Offline grace bez potpisanog dokaza | Nalaz | **Tačno** (:193–203; token se ne koristi u odluci) — dokumentovani plafon (:16–19) | P3 | Vidi 77.9 | — |
| 77.9 | `LICENSE_TOKEN` mrtav podatak | Nalaz | **Tačno** — jedini upis :416, nijedno čitanje (grep ceo src-vba) | P3 | Ili ukloniti ključ+komentar „potpisan token" (:35) ili ga stvarno verifikovati | S |
| 77.10 | Cache editabilan | Nalaz | **Dizajnersko ograničenje** — komentari priznaju (:16–19, :554–555) | Prihvaćeno | — | — |
| 77.11 | Fuzzy 2/3 collision/takeover površina | Nalaz | **Delimično** — klijent tačan (:66, :289–291); server-side policy praćenja **nije proverivo statički** ovde | P3 | Server-side beleška (GAS domen) | — |
| 77.12 | MachineGuid/serial/UUID nisu immutabilni | Nalaz | **Tačno** (kontekst; komentari :482–501 to i kažu) | — | — | — |
| 77.13 | Response nepotpisan; pin prazan | Nalaz | **Tačno** (:62 `LIC_ENDPOINT_PINNED=""`; :578–588 config override) | P3 | Vidi 77.14 | — |
| 77.14 | Pinning neaktivan | Nalaz | **Tačno** (:62); mehanizam spreman | P3 | Pri sledećem re-sign buildu upisati GAS URL u Const | S |
| 77.15 | Bez HTTPS/host validacije bez pina | Nalaz | **Tačno** (:447 endpoint direktno) | P3 | Minimalno: zahtevaj `https://` prefiks kad pin nije aktivan | S |
| 77.16 | Nepoznat status pušta vezanu mašinu | Nalaz | **Dizajnersko ograničenje** (:261–273, komentar N3) | Prihvaćeno | — | — |
| 77.17 | HTTP fail posle grace-a ipak pušta vezanu | Nalaz | **Tačno** + Dizajnersko (:211–219); `NEXT_CHECK` = rok re-provere, ne rok rada — FM ispravno opisuje | Prihvaćeno | Dokumentovati semantiku u runbook | S |
| 77.18 | Suspend/expiry neprimenjivi tokom outage-a | Nalaz | **Dizajnersko ograničenje** (posledica 77.17; FM sam kaže „očekivano") | Prihvaćeno | — | — |
| 77.19 | HWM dnevna rezolucija | Nalaz | **Tačno** (:189 `Date`, :307 `yyyy-mm-dd`) | P3 | — (dovoljno za kalendarsku odluku) | — |
| 77.20 | HWM persistence tih | Nalaz | **Tačno** (:303–309 `Resume Next`, upis neproveren) | P3 | LogWarn pri neuspehu upisa | S |
| 77.21 | Malformed HWM ignorisan | Nalaz | **Tačno** (:294–300 samo `IsDate` grana; :306 malformed se nikad ne prepiše) | P3 | Self-heal: malformed prepisati današnjim + log | S |
| 77.22 | Lokalno vreme, bez UTC | Nalaz | **Tačno** (:189, :418) | P3 | — | — |
| 77.23 | `PersistLicenseOk` neatomičan | Nalaz | **Tačno** (:411–420, 4 upisa bez provere) | P3 | Redosled: token/bound pre `NEXT_CHECK` (latch se ne aktivira polovično) | S |
| 77.24 | Aktivacija piše ključ/briše cache PRE servera | Nalaz | **Tačno** — :350–353 upis+brisanje pre HTTP :372; server fail (:374–377) ostavlja uništen stari validan cache, bez rollback-a. **Nova konkretna rupa (gubitak stanja legitimnog kupca), ne restatement fail-open-a** | **P2** | Staging: stari key/bound/next-check snimiti u lokalne promenljive, upisati novo TEK posle `status=OK`; na fail vratiti staro | S |
| 77.25 | Isti destructive pre-write u inline aktivaciji | Nalaz | **Tačno** (:327–329 pa `LicenseGateOrQuit`) | **P2** | Isti staging fix | S |
| 77.26 | Setter rezultati se ne proveravaju | Nalaz | **Delimično** — `SetConfigValue` je `Public Sub` bez povratne vrednosti (modConfig.bas:780); „ne proverava se" tačno, ali provera zahteva promenu potpisa | P3 | Uz 77.24 staging dovoljno | — |
| 77.27 | InputBox prikazuje postojeći ključ | Nalaz | **Tačno** (:346–347) — ali ključ je ionako plaintext u tblSEFConfig (plafon) | P3 | — | — |
| 77.28 | `LicenseShowDevice` public dijagnostika | Nalaz | **Tačno** + Dizajnersko (:397–405; namenjeno supportu) | P3 | — | — |
| 77.29 | Aktivacija bez admin guarda | Nalaz | **Tačno** (:341); reset cache-a = unutar plafona editabilnog configa | P3 | — | — |
| 77.30 | Nema rate limita | Nalaz | **Tačno**; odgovornost servera (GAS) | P3 | — | — |
| 77.31 | Nema retry/backoff | Nalaz | **Tačno** (:451 jedan Send) | P3 | — | — |
| 77.32 | HTTP rezultat samo Boolean | Nalaz | **Tačno** (:426–466) | P3 | — | — |
| 77.33 | Jednostavan JSON extractor | Nalaz | **Tačno** (:226 i dr.; server pod sopstvenom kontrolom) | P3 | — | — |
| 77.34 | `graceDays` permisivan, bez plafona | Nalaz | **Tačno** (:412–414: `CLng(val())`, ≤0→3, bez cap-a) | P3 | `If grace > 30 Then grace = 30` | S |
| 77.35 | `LICENSE_STATUS` se ne koristi | Nalaz | **Tačno** — jedini upis :419, bez čitanja (grep) | P3 | Ukloniti ili koristiti | S |
| 77.36 | Nema local deny cache-a | Nalaz | **Tačno** — SUSPENDED/EXPIRED (:237–243) ne dira cache; sledeći offline start = grace put. Editabilnost cache-a ga čini zaobilaznim (plafon), ali postoji jeftin delta | **P2** | Pri SUSPENDED/EXPIRED obrisati `BOUND_PARTS`+`NEXT_CHECK` (server je već rekao ne → offline grace prestaje) | S |
| 77.37 | Nema `LAST_CHECKED_AT`/policy verzije | Nalaz | **Tačno** | P3 | — | — |
| 77.38 | `gAccessDenied` bez reseta | Nalaz | **Tačno** — :78, :635; nigde `=False` (grep) | P3 | `gAccessDenied = False` na ulazu u `AccessGateOrQuit` | S |
| 77.39 | `OnTime` zatvaranje neprovereno | Nalaz | **Tačno** (:633–637 `Resume Next`, bez provere); ozbiljno tek u kombinaciji sa 78.27 | **P2** (klaster sa 78.27) | Posle `OnTime` proveriti `Err`, fallback direktan `ForceCloseDeniedWorkbook` | S |
| 77.40 | Force-close `Saved=True` odbacuje izmene | Nalaz | **Tačno** + Dizajnersko za startup (:642–647); gate se poziva samo iz StartApp | P3 | — | — |
| 77.41 | OnTime string nekvalifikovan | Nalaz | **Nije proverivo statički** (multi-workbook ponašanje `Application.OnTime` stringa; rizik plauzibilan) | P3 | `"'" & ThisWorkbook.Name & "'!modLicense.ForceCloseDeniedWorkbook"` | S |
| 77.42 | `LicenseBlock` guta greške | Nalaz | **Tačno** (:611–617) | P3 | Deo 77.39 fix-a | — |
| 77.43 | Bypass direktnim entry point-om | Nalaz | **Tačno — konkretan put potvrđen:** modPregledListova.bas:81–88 `PokreniProgram` (sheet dugme „Pokreni program") radi `frmOtkupAPP.Show` bez ikakve provere; komentar pogrešno pretpostavlja da je gate „već odrađen". Shift-open (preskočen `Workbook_Open`) ili pad OnTime close-a → ulaz bez gate-a, **bez VBE** — ispod dokumentovanog plafona | **P2** | U `PokreniProgram` dodati `If Not AccessGateOrQuit() Then Exit Sub` (fast offline put ga čini jeftinim) | S |
| 77.44 | Nema license audita | Nalaz | **Tačno** — 0 `Monitor_Event` u modulu | P3 | — | — |
| 77.45 | Raw fingerprint serveru | Nalaz | **Tačno** + Dizajnersko (server mora komponente za fuzzy 2/3) | P3 | Privacy/retention beleška u docs | S |
| 77.46 | 18 pozitivnih | Pozitivno | **Kontekst-Pozitivno** — potvrđeno | — | — | — |
| 77.47 | 24 hardening prioriteta | Prioriteti | **Tačno** (konzistentno); pun opseg = redesign van VBA plafona | P2 podskup | Usvojiti samo: staging aktivacije, deny-purge, graceDays cap, gAccessDenied reset, gate u `PokreniProgram` | S–M |

**Bilans:** 47/47 podsekcija verifikovano; ~18 = restatement dokumentovanog fail-open plafona (Prihvaćeno). **Stvarno nove akcione tačke ispod plafona:** destruktivna aktivacija bez rollback-a (77.24/25), izostanak deny-purge-a (77.36), potvrđen sheet-button bypass bez VBE (77.43) i OnTime klaster (77.39, sa 78.27). Sve su S popravke.

---

### FM-0076 — `modTrial.bas` (anchor a0bc9e2)

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 78.1 | Zaključak: deterrent, editabilan config, fail-open putevi; `Workbook_Open` bez `AccessWasDenied` | Zaključak | **Tačno** (cross-file deo potvrđen — vidi 78.27) | P1 (integracija) | Vidi 78.27 | S |
| 78.2 | Public API | Opis | **Tačno** (:46, :103, :119; privatni :137, :153, :162) | — | — | — |
| 78.3 | Orkestracija kroz `AccessGateOrQuit` | Opis | **Tačno** (modLicense.bas:103–115) | — | — | — |
| 78.4 | Config model + defaulti | Opis | **Tačno** (:27–29, :34–38, :41) | — | — | — |
| 78.5 | Fiksni kalendarski prozor, ne „N dana od prvog starta" | Nalaz | **Tačno** + Dizajnersko (:13 „zadati datum"; nigde first-run stamp) | P3 | Dokumentovati; first-run model samo ako se poslovno traži | M |
| 78.6 | Nema lower-bound `today<start` | Nalaz | **Tačno** (:60, :124 samo `today > deadline`) | P3 | Dodati `today >= start` ili dokumentovati pre-start ponašanje | S |
| 78.7 | Mogući off-by-one (11 datuma za `DAYS=10`) | Nalaz | **Tačno** kao dvosmislenost (:55 `start+days`, :60 `>` → uključivo start..start+days) | P3 | Definisati semantiku + acceptance primer u komentaru/testu | S |
| 78.8 | UI izlaže sve trial parametre | Nalaz | **Tačno** (modPodesavanja.bas:121–123; `TRIAL_HWM` sakriven :20) + Dizajnersko (operativni model po runbook-u) | P3 | — | — |
| 78.9 | `TRIAL_ENABLED=NO` + licenca off → pušta | Nalaz | **Tačno** (modLicense.bas:121–122) + Dizajnersko (runbook očekuje) | Prihvaćeno | — | — |
| 78.10 | Greška čitanja flaga tiho gasi trial | Nalaz | **Tačno** (:105–111 → default False :34, bez loga) | P3 | Log pri `Err` u čitanju | S |
| 78.11 | Parser bez `DA/NE` (uži od `ConfigFlag`) | Nalaz | **Tačno** (:108–112 vs modConfig.bas:890–892) — ublaženo: Podešavanja „bool" editor piše YES/NO (modPodesavanja:43) | P3 | Dodati `Case "DA"` / `Case "NE"` (ili reuse `ConfigFlag`) — isto i u `modLicense.LicenseEnabled` | S |
| 78.12 | `TRIAL_DAYS` permisivan parser | Nalaz | **Tačno** (:156–158 `CLng(val())`; ≤0→10) | P3 | — | — |
| 78.13 | Bez gornje granice → Date overflow → fail-open | Nalaz | **Tačno** (:55 → EH :93–97 True; `TrialActive` :120/:124 preskoči) — ali zahteva edit configa (plafon: isti korisnik može `TRIAL_ENABLED=NO`) | P3 | Cap npr. 3650 u `TrialDays` | S |
| 78.14 | `TRIAL_START` nije strict ISO | Nalaz | **Delimično** — `IsDate/CDate` locale-dependent tačno (:142–149); praktični rizik za `yyyy-mm-dd` nizak (isti obrazac koristi ceo projekat) | P3 | Ručni `yyyy-mm-dd` parse + `DateSerial` ako se ikad javi na terenu | S |
| 78.15 | Invalid start tiho pada na 18.06.2026 | Nalaz | **Tačno** (:142–149) | P3 | Log fallback-a | S |
| 78.16 | HWM deterrent koncept dobar | Pozitivno | **Kontekst-Pozitivno** (:66–88) | — | — | — |
| 78.17 | Malformed HWM trajno gasi anti-rollback | Nalaz | **Tačno** (:73–81 samo `IsDate` grana; :84 malformed se nikad ne prepiše) | P3 | Self-heal prepisom + log (malformed nastaje samo ručnim editom = plafon) | S |
| 78.18 | HWM round-trip zavisi od locale-a | Nalaz | **Delimično** (piše :86 ISO, čita :74 `IsDate` — teorijski jaz, praktično stabilan na Windows VBA) | P3 | — (isti fix kao 78.14 ako zatreba) | — |
| 78.19 | HWM upis potpuno tih | Nalaz | **Tačno** (:85–87 `Resume Next`; + perzistencija zavisi i od save-a sveske) | P3 | Log pri neuspehu | S |
| 78.20 | HWM dnevna rezolucija | Nalaz | **Tačno** (:56, :86) + Dizajnersko za kalendarski trial | P3 | — | — |
| 78.21 | Skok sata unapred = trajni lockout | Nalaz | **Tačno** (:73–88); runbook workaround postoji (production-runbook-licenca.md:133 — obriši `TRIAL_HWM`) | P3 | Po potrebi admin reset makro | S |
| 78.22 | `TrialActive`+`TrialGateOrQuit` dupliraju odluku | Nalaz | **Tačno** (modLicense.bas:104–106; obe čitaju config/datum) — posledice minimalne (isti trenutak) | P3 | `EvaluateTrial` refaktor → u refactor paket, ne hitno | M |
| 78.23 | `TrialActive` globalni `Resume Next` | Nalaz | **Tačno** (:120; ishod zavisi od tačke greške) | P3 | Uz 78.22 | — |
| 78.24 | Gate eksplicitno fail-open na exception | Nalaz | **Dizajnersko ograničenje** (:93–97, komentar) — nijansa da EH pokriva i overflow pre odluke je tačna | Prihvaćeno/P3 | Cap iz 78.13 uklanja overflow slučaj | S |
| 78.25 | Komentar „bez zavisnosti" preširok | Nalaz | **Tačno** (:58–59 vs :55 — deadline zavisi od 2× `GetConfigValue` + Date opsega) | P3 | Ispraviti komentar | S |
| 78.26 | Blokada zavisi od neproverenog `OnTime` | Nalaz | **Tačno** (:162–172 → modLicense.bas:633–637) | **P2** (klaster) | Fix u `DenyAccessAndScheduleClose` (vidi 77.39) | S |
| 78.27 | **Kritično cross-file: `Workbook_Open` ne čita `AccessWasDenied`** | Kritično | **Tačno — potvrđeno:** ThisWorkbook.doccls:15–35 nema provere posle `StartApp`; `AccessWasDenied` (modLicense.bas:626–628) **nigde pozvan** (grep = 0 poziva u kodu); a modLicense.bas:75–77 komentar i docs/production-runbook-licenca.md:86 tvrde da provera postoji. Dokumentacija ≠ kod | **P1** | U `Workbook_Open`, odmah posle `StartApp`: `If AccessWasDenied() Then Exit Sub` | S |
| 78.28 | Lažni `VBA_STARTUP_SUCCESS` posle deny-ja | Nalaz | **Tačno** (ThisWorkbook.doccls:24–35 bezuslovno posle StartApp) | **P1** | Isti fix kao 78.27 (early-exit pre Monitor_Event) | S |
| 78.29 | Gate nije prva startup granica | Nalaz | **Tačno** (modMain.bas:19–38: Monitor→`InitApp` sa EnsurePoruke:188/EnsureRuntimeSchema:198/ValidateAllTables:205 pre gate-a :38) + Dizajnersko pitanje | P3 | Dokumentovati nameru („blokira UI, ne init") | S |
| 78.30 | Direktni makroi zaobilaze gate | Nalaz | **Tačno** — isti put kao 77.43 (modPregledListova.bas:83) | **P2** | Fix iz 77.43 | S |
| 78.31 | Nema typed rezultata | Nalaz | **Tačno** (Boolean API) | P3 | — | — |
| 78.32 | Nema trial audita | Nalaz | **Tačno** (samo `LogErr` :96) | P3 | — | — |
| 78.33 | Testovi ne pokrivaju trial | Nalaz | **Tačno** (modLicenseTests = samo fingerprint core) | P3 | Par čistih testova za deadline granice uz 78.7 | S |
| 78.34 | 15 pozitivnih | Pozitivno | **Kontekst-Pozitivno** — potvrđeno (uklj. HWM van editora) | — | — | — |
| 78.35 | 22 prioriteta | Prioriteti | **Tačno**; #1–2 su pravi hitni | P1 (#1–2), P2 (#3), P3 ostalo | Usvojiti #1–2 odmah; #3 uz 77.39 | S |

**Bilans:** najjači entry paketa — **78.27/78.28 su potvrđene stvarne integracione greške P1** (kod protivreči sopstvenom komentaru i runbook-u; deny se oslanja isključivo na neproveren `OnTime`), fix je 2 linije. Ostalo: pretežno tačno opisan deterrent dizajn (Prihvaćeno/P3) + par S poliranja (DA/NE parser, off-by-one semantika, HWM self-heal).

---

### FM-0077 — `modLicenseTests.bas` (anchor a0bc9e2)

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 80.1 | Zaključak: koristan core suite, ime šire od obuhvata, duplikat postoji, ranije brisan pa vraćen | Zaključak | **Delimično** — sve potvrđeno osim „ranije brisan" (istorija neproveriva — shallow klon; vidi 80.33) | P1 (duplikat → RF-01) | RF-01 | S |
| 80.2 | Deklarisani scope pošten; server-side `runLicenseSelfTest` | Opis | **Tačno** (:4–13; gas/Code.gs:6029) | — | — | — |
| 80.3 | Javni API + nema `Option Private Module` | Opis | **Tačno** (:18–90 Public; :96/:106 Private; :15–16) | — | — | — |
| 80.4 | Model izvršavanja bez povratne vrednosti/gate-a | Opis | **Tačno** (:18–33) | — | — | — |
| 80.5 | Tačno 23 asercije, broj se ne proverava | Opis | **Tačno** — prebrojano 8+9+5+1=23; nema expected-count | P3 | Uz 80.24 | S |
| 80.6 | SplitParts pokrivenost | Pozitivno | **Kontekst-Pozitivno** (:40–52) | — | — | — |
| 80.7 | PartsMatch pokrivenost | Pozitivno | **Kontekst-Pozitivno** (:59–69, svih 9 potvrđeno) | — | — | — |
| 80.8 | NonEmptyParts pokrivenost | Pozitivno | **Kontekst-Pozitivno** (:76–80) | — | — | — |
| 80.9 | Smoke zove pravi `GetDeviceParts` | Pozitivno | **Kontekst-Pozitivno** (:84–90) | — | — | — |
| 80.10 | 2/3 granica dobro pogođena | Pozitivno | **Kontekst-Pozitivno** | — | — | — |
| 80.11 | Prag 2 dupliran, ne čita `LIC_MIN_MATCH` | Nalaz | **Tačno** — Const je `Private` (modLicense.bas:66); test hardkod (:68–69, :80) | P3 | `Public Function LicMinMatch()` (ili Public Const) + test je čita | S |
| 80.12 | `LicenseIsBoundMachine` netestiran | Nalaz | **Tačno** (Private, modLicense.bas:289–291) | P3 | Pure `EvaluateDeviceMatch` seam | M |
| 80.13 | `TestLicense_All` preširoko ime | Nalaz | **Tačno** — runbook ga koristi kao glavni suite (production-runbook-licenca.md:88) | P3 | Napomena o obuhvatu u header print ili rename | S |
| 80.14 | Nema gate matrix testova | Nalaz | **Tačno** | P3 | — | M |
| 80.15 | Nema contract testova | Nalaz | **Tačno** | P3 | — | M |
| 80.16 | Nema offline-proof testova | Nalaz | **Tačno** | P3 | — | M |
| 80.17 | Nema activation/persistence testova | Nalaz | **Tačno** (najrizičniji mutation tok — poklapa se sa 77.24) | P3 | Uz staging fix 77.24 dodati 2–3 testa | M |
| 80.18 | Env smoke pomešan sa unit | Nalaz | **Tačno** (:84–90 u istom runneru) | P3 | Odvojen entry point | S |
| 80.19 | Smoke meri količinu, ne kvalitet | Nalaz | **Tačno** (jedina asercija :89) | P3 | — | — |
| 80.20 | Ispis kompletnog raw otiska | Nalaz | **Tačno** (:88 `Debug.Print "Otisak: "`) | P3 | Maskirati (prva 4 znaka + `...`) | S |
| 80.21 | Nema component dijagnostike | Nalaz | **Tačno** (Read* gutaju greške — modLicense.bas:484,492,503) | P3 | Per-komponenta DA/NE ispis | S |
| 80.22 | Jedna runtime greška obara suite | Nalaz | **Tačno** (:18–33 bez handlera) | P3 | — | S |
| 80.23 | Fail ne obara release/automatizaciju | Nalaz | **Tačno** (:96–114 samo brojači) | P3 | Za ručni Alt+F8 tok dovoljan summary; `Err.Raise` tek uz automatizaciju | S |
| 80.24 | False-green `PASS=0 FAIL=0` | Nalaz | **Tačno** (:30–31 samo `mFail=0`) | P3 | `Const EXPECTED=23` + provera `mPass+mFail` | S |
| 80.25 | Pojedinačni testovi bez lifecycle-a | Nalaz | **Tačno** (samo `TestLicense_All` resetuje :19) | P3 | Pojedinačne učiniti Private | S |
| 80.26 | Assert helperi nisu strict | Nalaz | **Tačno** (Variant + `=`) | P3 | — | — |
| 80.27 | Fale boundary/malformed slučajevi | Nalaz | **Tačno** | P3 | Dodati `A|B`, `A|B|C|D`, whitespace slučajeve | S |
| 80.28 | SplitParts „bar 3", ne fixed-width | Nalaz | **Tačno** (modLicense.bas:518–519 `Split(s & "||")` → 5 elemenata za pun otisak; test samo `UBound>=2` :45) | P3 | — (produkcija koristi 0..2; dokumentovati) | — |
| 80.29 | Nema table-driven pokrivenosti | Nalaz | **Tačno** | P3 | — | S |
| 80.30 | **Kritično: postoji `modLicenceTests.bas`** | Kritično | **Tačno** — src-vba/modLicenceTests.bas:1–2 (`Attribute VB_Name = "modLicenceTests"` + `'Attribute VB_Name = "modLicenseTests"`); telo identično kanonskom (diff = samo header + trailing newline); istih 5 Public procedura | **P1** — **rešenje već planirano: RF-01 (brisanje)** | Sprovesti RF-01 | S |
| 80.31 | `ImportAllVBA` uvozi oba | Nalaz | **Tačno** — modVbaTools.bas:80–88: Remove samo istog baseName, Import po header `VB_Name`; nema manifest/dup/exclusion/compile gate-a | **P2** (nadživljava RF-01) | Pre-import/pre-release dup-validator (vidi 82.16) | M |
| 80.32 | Posledice duplog modula | Nalaz | **Tačno** (dobro ograđeno) — unqualified poziva u kodu nema pa compile ostaje čist; Alt+F8 dupli entry realan | P1 (kroz RF-01) | RF-01 | — |
| 80.33 | Duplikat ranije poznat i ponovo uveden | Nalaz | **Nije proverivo statički** — klon je shallow (8 graft root-ova); citirana commit poruka („Obrisani sveskini dupli/typo test-moduli…") ne postoji u dostupnoj istoriji; vidljivo samo dodavanje kroz merge PR #92 (`7091213`, grana `vba-workbook-migration-modules` — konzistentno sa workbook-export poreklom) | — | Validator iz 80.31/82.16 štiti nezavisno od istorije | — |
| 80.34 | Test makroi u produkcionoj macro površini | Nalaz | **Tačno** (nema `Option Private Module`) | P3 | `Option Private Module` u test module | S |
| 80.35 | Produkcioni helperi Public radi testa | Nalaz | **Tačno** + Dizajnersko (modLicense.bas:469–473, :512–514 to dokumentuju) | P3 | — | — |
| 80.36 | Nema build identiteta u outputu | Nalaz | **Tačno** (:21 samo `Now`) | P3 | Header print `APP_VERSION`/`BUILD_SHA` + module name (rešava i 82.9) | S |
| 80.37 | Suite read-only | Pozitivno | **Kontekst-Pozitivno** | — | — | — |
| 80.38 | 14 pozitivnih | Pozitivno | **Kontekst-Pozitivno** — potvrđeno | — | — | — |
| 80.39 | 26 prioriteta | Prioriteti | **Tačno**; #1=RF-01 (planiran), #2–4 validator | P1 (#1, planiran); P2 (#2–4); P3 ostalo | RF-01 + dup-validator; od ostalog samo build-header (#25) i expected-count (#9) | S–M |

**Bilans:** 39/39 verifikovano; suite radi tačno ono što tvrdi u komentaru, a FM-ove kritike obuhvata su tačne ali P3 (ručni dev alat). Jedina korekcija: tvrdnja o **ranijem brisanju nije dokaziva iz shallow klona**. Akciono: RF-01 (planiran) + trajni dup-validator (P2) + sitni S dodaci (build header, expected count, maskiran otisak).

---

### FM-0078 — `modLicenceTests.bas` (anchor a0bc9e2; **fajl ide u RF-01 brisanje**)

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 82.1 | Duplikat bez sopstvene funkcije; kontradiktorni header | Zaključak | **Tačno** (:1–2; telo identično) — nalaz = obrazloženje RF-01 | P1 → RF-01 | Sprovesti RF-01 | S |
| 82.2 | Razlika prema kanonskom = samo identitet | Nalaz | **Tačno** — diff: linija 1–2 headera + završni newline; ostalo bajt-identično | — | — | — |
| 82.3 | Duplira 5 Public procedura | Nalaz | **Tačno**; bespredmetno posle RF-01 (fajl se briše) | RF-01 | — | — |
| 82.4 | `ImportAllVBA` uvozi oba; importer bez zaštita | Nalaz | **Tačno** (modVbaTools.bas:80–88) — tooling deo **preživljava RF-01** | **P2** | Dup-validator pre import/release (82.16) | M |
| 82.5 | Repo istorija potvrđuje raniji delete | Nalaz | **Nije proverivo statički** (shallow klon; kao 80.33) | — | — | — |
| 82.6 | Provenance: workbook-export artefakt | Nalaz | **Delimično** — eksplicitno označen inference; header + naziv PR grane (`vba-workbook-migration-modules`) konzistentni | — | — | — |
| 82.7 | Nema compatibility vrednost (nije wrapper) | Nalaz | **Tačno** (pun copy); bespredmetno posle RF-01 | RF-01 | — | — |
| 82.8 | Drift sada 0, ali gotovo zagarantovan | Nalaz | **Tačno** (trenutno identični); bespredmetno posle RF-01 | RF-01 | — | — |
| 82.9 | Umanjuje dokaznu vrednost `TestLicense_All` rezultata | Nalaz | **Tačno** (runbook:88; ni jedan suite ne ispisuje module/build) — posle RF-01 ostaje samo build-id deo = FM-0077/80.36 | RF-01 + P3 | Build/module header u kanonskom suite-u | S |
| 82.10 | Širi produkcionu macro površinu | Nalaz | **Tačno**; bespredmetno posle RF-01 | RF-01 | — | — |
| 82.11 | Nasleđuje sve slabosti kanonskog | Nalaz | **Tačno** po identičnosti; bespredmetno posle RF-01 (praćeno pod FM-0077) | RF-01 | — | — |
| 82.12 | Nema zavisnosti/domain ownership | Nalaz | **Tačno**; bespredmetno posle RF-01 | RF-01 | — | — |
| 82.13 | Rizik po poslovne podatke nizak | Kontekst | **Tačno** (read-only) | — | — | — |
| 82.14 | Import/compile rizik visok (kanonski folder) | Nalaz | **Tačno** — leži tačno na putanji koju `ImportAllVBA` enumeriše (modVbaTools.bas:44–45) | P1 do RF-01 | RF-01 | S |
| 82.15 | Minimalna ispravka (6 koraka) | Predlog | **Tačno** — koraci = tačno RF-01 (brisanje, uklanjanje komponente, Compile, kvalifikovan test, re-export provera); ništa za merge | P1 | Usvojiti kako piše | S |
| 82.16 | Trajna zaštita u tooling-u (validator) | Predlog | **Tačno/opravdano** — **preživljava RF-01**; danas ne postoji nijedna provera (tools/ = samo release/stamp skripte) | **P2** | Skript u `tools/`: dup `VB_Name`, filename↔VB_Name, normalizovan spelling, dup Public Sub/Function; pozvati iz release.sh | M |
| 82.17 | Manifest-driven import | Predlog | **Tačno/opravdano**, ali teži zahvat za dev alat | P3 | Validator iz 82.16 je jeftiniji prvi korak | M |
| 82.18 | Export ne prepoznaje rename/delete | Nalaz | **Tačno** — `ExportAllVBA` (modVbaTools.bas:21–42) izvozi preko postojećeg foldera; nema staging-a ni brisanja stale fajlova → tačno klasa greške koja je vratila duplikat | **P2** | Export u prazan staging folder + poređenje/brisanje viška (ili bar uputstvo u komentaru) | S–M |
| 82.19 | Release gate da odbije ovaj repo state | Predlog | **Tačno/opravdano** — deo istog validatora | P2 | U sklopu 82.16 | — |
| 82.20 | Lako i bezbedno ukloniti | Pozitivno | **Kontekst-Pozitivno** — potvrđeno: 0 referenci na `modLicenceTests` u kodu/tools (grep), kanonski identičan | — | — | — |
| 82.21 | 18 prioriteta | Prioriteti | **Tačno**; #1–5=RF-01, #6–13 tooling, #14–16=FM-0077 teren | P1/P2/P3 | RF-01 odmah; validator+čist export kao poseban mali PR | S+M |
| 82.22 | Funkcionalni zaključak (0 funkcija, 5 duplih entry-ja) | Zaključak | **Tačno** | RF-01 | — | — |

**Bilans:** sve provereno tačno u sadašnjem stanju; jedino istorijski deo (82.5, delom 82.6) nije proveriv iz shallow klona. Fajl je čist artefakt — po instrukciji: file-scoped nalazi su **bespredmetni posle RF-01**; trajnu vrednost entry-ja nose 82.4/82.16/82.18/82.19 (dup-validator + staging export, P2 M), koje treba preneti u backlog nezavisno od brisanja fajla.

---

**Zbirno preko svih 6 entry-ja (151 podsekcija):** dominira **Tačno/Dizajnersko** — FM činjenično veoma precizno čita kod. Korekcije: 2× „Nije proverivo statički" (git istorija — shallow klon: 80.33, 82.5), nekoliko **Delimično** (75.2-nr.30, 77.11, 77.26, 77.41, 78.14, 78.18, 82.6). **Pravi hitni ostatak posle kalibracije:** (P1) `Workbook_Open` bez `AccessWasDenied` + lažni STARTUP_SUCCESS (78.27/78.28, fix 2 linije) i RF-01 brisanje duplikata (planirano); (P2) staging aktivacije licence (77.24/25), deny-purge (77.36), `PokreniProgram` gate (77.43/78.30), OnTime provera (77.39/78.26), dup-validator + staging export u tooling-u (80.31/82.16/82.18).

---

## Delta blok 9 — Build guard/info, business-flow testovi, cenovnik, config re-audit, E2E gate (FM-0079…FM-0084, 187 podsekcija) [sidro a0bc9e2]

Sva verifikacija je završena. Slede kompletne tabele.

# Audit FM-0079 … FM-0084 protiv koda na `a0bc9e2` (`/home/user/otkupapp-pwa/src-vba/`)

Napomena o kalibraciji: single-writer, ručni release proces; pre-registrovani nalazi (AUD-003, AUD-016, AUD-018, KI-006, RF-01) se referenciraju, ne eskaliraju ponovo.

### FM-0079 — `modBuildGuard.bas` (63 linije; kod verifikovan u celini)

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 84.1 | Guard proverava samo `ListObject` redove; nije blanko-garancija (FP na seed, FN van tabela) | Sinteza | **Tačno** (petlja samo `ws.ListObjects`, modBuildGuard.bas:29-32) | P2 | Vidi 84.7/84.5 | — |
| 84.2 | Stvarni algoritam = broj fizičkih redova po ListObject-u | Kontekst | **Tačno** (modBuildGuard.bas:29-41) | — | — | — |
| 84.3 | `Alt+F8` pokazuje na `Public Function`, ne runnable Sub | Rizik | **Delimično** — Function se ne LISTA u Macro dijalogu, ali se izvršava upisom imena; tvrdnja „mora preko Immediate/druge Sub" je prejaka | P3 | Tanki `Public Sub RunBlankBuildCheck` wrapper | S |
| 84.4 | Binarni kriterijum protivreči poruci o dozvoljenim seed šifarnicima | Rizik | **Tačno** (poruka modBuildGuard.bas:52-54 priznaje seed, a rezultat je `False`) | P3 | Allowlist seed tabela u guardu | S |
| 84.5 | Normalan startup puni `tblPoruke` → guard uvek crven | Rizik | **Tačno** (Workbook_Open→StartApp→InitApp→`EnsurePoruke`, modMain.bas:188; ~226 `UpsertRow` u modPoruke) — ali RELEASE_PROCEDURE.md:223 dokumentuje „isprazni pa ponovi" | P3 | `tblPoruke` tretirati kao poznati seed (izuzetak + napomena u reportu) | S |
| 84.6 | Broji fizičke redove, ne sadržaj (FP na prazan red/formule) | Rizik | **Tačno** (`DataBodyRange.rows.count`, :32); FP je safe-side | P3 | Ništa hitno; opciono `CountA` napomena | S |
| 84.7 | **FN: `SETUP_LOG`** (mašina/korisnik/putanja) nevidljiv guardu | Kritično | **Tačno** — guard skenira samo `ws.ListObjects` (modBuildGuard.bas:30), a `SETUP_LOG` je plain range (modSetup.bas:28, InitSetupLog modSetup.bas:1729-1741) u koji se upisuju `ThisWorkbook.fullName`, `COMPUTERNAME`, `USERNAME`, verzija Excela (modSetup.bas:52-55) | **P2** | U `AssertBlankBuild` dodati proveru poznatih plain-range logova (`SETUP_LOG`, `BUSINESS_FLOW_PRO_TEST_LOG`, `NOVAC_TEST_LOG`…): sheet postoji + ima >1 red → prijavi | S |
| 84.8 | Ostale neproverene površine (named ranges, hidden ćelije, properties…) | Rizik | **Tačno** (trivijalno iz koda) | P3 | Dokumentovati scope u header komentaru | S |
| 84.9 | Nema data-policy klasifikaciju (REQUIRED_EMPTY/SEED…) | Predlog | **Tačno** kao činjenica; puni registry je overkill za obim | P3 | Mini-allowlist iz 84.4/84.5 dovoljan | S |
| 84.10 | Workbook bez ijednog ListObject-a = „BLANKO OK" | Rizik | **Tačno** (`nonEmpty=0`→True, :41-45) | P3 | Upozorenje ako je nađeno 0 tabela | S |
| 84.11 | Ne proverava build stamp (placeholder/dirty) | Rizik | **Tačno** (nema reference na `BUILD_*` u modulu; placeholder potvrđen modBuildInfo.bas:5-7) | P2 (u paketu 86.3/86.8) | Deny placeholder/`+dirty` u publish koraku, ne nužno ovde | S |
| 84.12 | `tools/release.sh` samo ispisuje instrukciju; `PublishReleaseToDrive` ništa ne proverava | Rizik | **Tačno** (tools/release.sh:61 samo echo; modRelease.bas:22-66 bez ikakvog guarda) | **P2** | U `PublishReleaseToDrive`: abort ako `BUILD_SHA="0000000"` ili sadrži `+dirty` | S |
| 84.13 | TOCTOU: provera nije vezana za artefakt | Rizik | **Dizajnersko ograničenje** ručne procedure (bash ne može da pokrene Excel korak) | P3 | Ponoviti guard kao deo Save As uputstva | S |
| 84.14 | Rezultat je samo Boolean + MsgBox | Rizik | **Tačno** (:21, :41-56) | P3 | Vratiti report string ByRef po potrebi | S |
| 84.15 | MsgBox report nije skalabilan | Rizik | **Tačno**; realan broj tabela je mali | P3 | `Debug.Print` full report uz MsgBox rezime | S |
| 84.16 | „Build-only" granica samo u komentaru (nema `Option Private Module`) | Rizik | **Tačno** (nijedan .bas u repou nema `Option Private Module`); modul je read-only pa je rizik mali | P3 | `Option Private Module` (pazi: Application.Run/UDF posledice) | S |
| 84.17 | Error put fail-closed, ali generičan i neprocesan | Rizik | **Tačno** (:59-62) | P3 | — | S |
| 84.18 | Nema automatskih testova guarda | Rizik | **Tačno** (`AssertBlankBuild` se ne poziva ni iz jednog test modula ni E2E gate-a) | P3 | Skupo u VBA; preskočiti | M |
| 84.19 | Dobre osobine (mali, ThisWorkbook, nedestruktivan, hidden sheetovi pokriveni…) | Pozitivno | **Kontekst-Pozitivno** — sve tvrdnje potvrđene | — | — | — |
| 84.20 | Ciljni dizajn (policy registry, receipt, BuildReleaseArtifact) | Predlog | **Tačno** u premisama; predimenzionirano za jednog operatera | — | Uzeti samo 84.7 + 84.12 delove | L |
| 84.21 | 25 hardening prioriteta | Predlog | **Tačno** u premisama; lista meša S i L zahvate | — | Minimalni paket: stavke 1, 7, 12, 13 | — |
| 84.22 | Zaključak: koristan smoke check, ne pouzdana kapija | Sinteza | **Tačno** — fer profil | — | — | — |

**Bilans FM-0079:** 22 podsekcije: 17 Tačno, 1 Delimično (84.3), 1 Dizajnersko ograničenje (84.13), 1 Kontekst-Pozitivno, 2 sinteze/predloga bez zamerki. Nema Netačno. Realno eskalira: **84.7 (SETUP_LOG FN, P2)** i **84.12 (publish bez ikakvog guarda, P2)** — oba pokriva jedan S-patch (scan plain-range logova + placeholder/dirty deny u `PublishReleaseToDrive`). Ostalo P3.

### FM-0080 — `modBuildInfo.bas` (7 linija; auto-generisan — pre-registrovan kontekst)

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 86.1 | Pasivni nosač 3 konstante; nema validacije/vezivanja | Sinteza | **Tačno** (modBuildInfo.bas:5-7; nigde validator) | — | — | — |
| 86.2 | Veliki blast radius (monitoring/licenca/update/release) | Kontekst | **Tačno** (modMonitoring.bas:398-400, modLicense.bas:440-442, modUpdateGate.bas:99-100, modRelease.bas:71-77) | P3 | — | — |
| 86.3 | Placeholder je validan runtime identitet; nigde se ne odbija | Rizik | **Tačno** (grep: nema `BuildIdentityStatus`/placeholder provere) | **P2** | Isti S-patch kao 84.12: deny u `PublishReleaseToDrive` | S |
| 86.4 | `BUILD_DATE` = committer date, ne datum builda | Rizik | **Tačno** (`git show -s --format=%cI HEAD`, tools/stamp-build.sh:10) | P3 | Preimenovati/komentar u generatoru | S |
| 86.5 | Nema identitet konkretnog `.xlsm` artefakta | Rizik | **Tačno**; hash `.xlsm` iz VBA je nategnut | P3 | Hash korak u release checklisti (spoljni alat) | M |
| 86.6 | Short SHA bez full SHA | Rizik | **Tačno** (tools/stamp-build.sh:9) | P3 | Dodati full SHA u version.json | S |
| 86.7 | `git describe --tags` bez `--match "vba-v*"` | Rizik | **Tačno** (stamp-build.sh:13); danas u repou nema konkurentskih tagova → prospektivan rizik | P3 | `--match "vba-v*"` u oba skripta | S |
| 86.8 | Dirty se označava, ne zabranjuje | Rizik | **Tačno** za navedene module; napomena: `release.sh:24-28` odbija prljav radni dir na happy-path-u (delimična mitigacija koju FM ne pominje) | **P2** | Deny `+dirty` u `PublishReleaseToDrive` (isti patch) | S |
| 86.9 | `APP_VERSION` ↔ `BUILD_VERSION` bez validacije para | Rizik | **Tačno** (APP_VERSION modConfig.bas:13; nigde poređenje) | P3 | U publish-guardu: `BUILD_VERSION` base = `vba-v` & APP_VERSION | S |
| 86.10 | Monitoring prenosi identitet, ali mu slepo veruje | Rizik | **Tačno** | P3 | — | — |
| 86.11 | Build podaci nisu trust signal (lokalno editabilni) | Rizik | **Dizajnersko ograničenje** VBA platforme; tačno opisano | Prihvaćeno | — | — |
| 86.12 | Update gate odlučuje samo po `APP_VERSION >= minVersion` | Rizik | **Tačno** (modUpdateGate.bas:50) | P3 | Server-side revocation = feature, ne bug | M |
| 86.13 | Manifest (workbook konstante) vs. upload (disk `SRC_FOLDER`) bez cross-checka | Rizik (visok) | **Tačno** — modRelease.bas:19 (hardkodovan `C:\Users\Dusan\...`), :38-48 (upload sa diska), :71-77 (manifest iz konstanti workbooka); nikakve provere podudarnosti | **P2** | Pre uploada parsirati `BUILD_SHA` iz disk `modBuildInfo.bas` i uporediti sa workbook konstantom; abort na mismatch | S |
| 86.14 | Self-update može primeniti nov identitet na parcijalan kod | Rizik (visok) | **Tačno** — skip lista je samo `modSelfUpdate`+`modVbaTools` (modSelfUpdate.bas:33); faza 2 snima workbook i uz `stillFail` (modSelfUpdate.bas:171-180) | P2 | `modBuildInfo` primeniti poslednji + ne snimati „uspešno" uz stillFail | M |
| 86.15 | Obrnuto: nov kod + star identitet (nema inventory/hash manifesta) | Rizik | **Tačno** (`DownloadReleaseFiles` samo broji preuzeto, modSelfUpdate.bas:261-281) | P2 (isti paket) | Minimalno: lista očekivanih fajlova u version.json + provera pre primene | M |
| 86.16 | Nema PENDING/APPLIED transaction marker | Rizik | **Tačno** | P3 (pokriva 86.14/15 paket) | — | M |
| 86.17 | Repo namerno ne čuva shipped stamp; git nedovoljan za „šta je klijent dobio" | Rizik | **Tačno** (release.sh: `git checkout -- src-vba/modBuildInfo.bas`); tag+version.json daju delimičnu sledljivost | P3 | Redak release-log zapis u RELEASE_NOTES | S |
| 86.18 | Nema channel/schema/compat identitet | Rizik | **Tačno**; jedan kanal je realnost projekta | P3 | — | M |
| 86.19 | Raw konstante bez typed accessor-a | Rizik | **Tačno** | P3 | — | M |
| 86.20 | Bash/PS1 usklađeni, ali duplirana logika bez parity testa | Rizik | **Tačno** (obe skripte pročitane — identična logika) | P3 | Komentar-upozorenje o sinhronizaciji | S |
| 86.21 | Nema testova/CI za stamp | Rizik | **Tačno** (nema `.github/workflows` ni `tests/`) | P3 | — | M |
| 86.22 | Dobro rešeno (centralno, auto iz gita, dirty oznaka, ASCII…) | Pozitivno | **Kontekst-Pozitivno** — potvrđeno | — | — | — |
| 86.23 | Ciljni dizajn (manifest/receipt/tranzakcija) | Predlog | Premise tačne; obim L | — | — | L |
| 86.24 | 25 prioriteta | Predlog | Tačne premise; minimalni paket = 2, 3, 6, 11 | — | — | — |
| 86.25 | Zaključak: dobar observability temelj, nije dokaz code seta | Sinteza | **Tačno** | — | — | — |

**Bilans FM-0080:** 25 podsekcija: 20 Tačno (uz jednu izostavljenu mitigaciju u 86.8 — `release.sh` clean-check), 1 Dizajnersko ograničenje (86.11), 1 Kontekst-Pozitivno, 3 sinteze/predlozi. Nema Netačno. Eskalacije: **86.3+86.8 (placeholder/dirty deny — isti S-patch kao 84.12)** i **86.13 (disk↔workbook cross-check, S)**; 86.14/86.15 su realni M-zahvati u `modSelfUpdate` (redosled primene identiteta).

### FM-0081 — `modBusinessFlowProTests.bas` (2388 linija; kod pročitan u celini)

Napomena kalibracije: pre-registrovani opis „runs inside always-rollback TX" **ne važi za ovaj modul** — u fajlu nema nijednog `clsTransaction`/`BeginTx` (grep = 0); rollback TX koristi samo `RunMasterSyncSmokeSuite` (modGoogleSyncSmokeTests.bas:364-379). FM-ova tvrdnja 87.7 je dakle ispravna, a kontekstna beleška se odnosi na drugi suite.

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 87.1 | Vredna domenska suite, ali 4 pomešane uloge bez izolacije | Sinteza | **Tačno** | — | — | — |
| 87.2 | Testira stvarne produkcione `*_TX` ulaze; jak full-chain | Pozitivno | **Kontekst-Pozitivno** (svi navedeni pozivi potvrđeni, npr. :245, :362-471) | — | — | — |
| 87.3 | „Empty workbook" je samo komentar; nema preflight guarda | Rizik | **Tačno** (RunBusinessFlowProSuite :60-86 odmah seeduje/mutira) | **P2** | Guard na početku: ako `tblOtkup` ima ne-`TST-PRO` redove → traži potvrdu/odbij | S |
| 87.4 | Test modul se isporučuje klijentima; javni destruktivni makroi | Rizik | **Tačno** (PublishReleaseToDrive šalje sve `.bas` iz src-vba, modRelease.bas:38-48; nema `Option Private Module`) | **P2** | Guard iz 87.3 (izbacivanje iz manifesta kvari self-update starih kopija → M) | S–M |
| 87.5 | `AutoLinkOtkupOtpremnica_TX` bez scope-a može povezati realne redove | Rizik (visok) | **Tačno** (modSledljivost.bas:85+ radi nad celim tabelama; test :442, :1173, :1238 poziva bez filtera) | P2 | Uz 87.3 guard dovoljno; ili opcioni parametar allowed-IDs | M |
| 87.6 | Suite proizvodi realne monitoring događaje bez isTest | Rizik | **Tačno** (OTKUP_SAVE_SUCCESS modOtkup.bas:43, OTKUP_MULTI... :303, FAKTURA_CREATE... modFaktura.bas:35, SLEDLJIVOST_AUTOLINK... modSledljivost.bas:30) | P3 | `isTest` polje u payload — širi zahvat | M |
| 87.7 | Nema suite-level transakcije/rollbacka | Rizik | **Tačno** (0 TX u modulu) | P3 | Disposable kopija workbooka u proceduri testiranja | S (proces) |
| 87.8 | Fiksni seed ID-jevi se proveravaju samo po postojanju | Rizik | **Tačno** (`If RowExists(...) Then Exit Sub`, :1418 itd.) | P3 | Fingerprint provera ključnih polja | S |
| 87.9 | `ApplyAvansToFaktura` na fiksnom TEST_KUP_ID | Rizik | **Tačno** (modFaktura.bas:331) | P3 | — | — |
| 87.10 | Seed direktno kroz `AppendRow` | Rizik | **Tačno** (:1694-1698); prihvatljivo za fixture | Dizajnersko ograničenje | — | — |
| 87.11 | Config izmene nisu crash-safe | Rizik | **Tačno** (:866-893, :1467-1519 — restore na normal+EH putu, ne na crash) | P3 | Startup marker overkill; dokumentovati | S |
| 87.12 | Fixture 2090 vs. faktura `Date` | Rizik | **Tačno** (NextTestDate :1940-1943; fakturaRow `Date` modFaktura.bas:~264) | P3 | Test assert za datum fakture | S |
| 87.13 | Negativni testovi prihvataju bilo koju grešku | Rizik | **Tačno** (:588-664 `ExpectedError` samo count; :554-564 dup-faktura svaki `Err.Number<>0` = PASS) | P3 | Proveriti `Err.Number` opseg | S |
| 87.14 | Fail-soft helperi mogu dati false-green | Rizik | **Tačno** (CountRows/RowExists/GetValueByKey :1700-1752 vraćaju 0/False/Empty na grešku) | P3 | Hard-fail verzije za asertacije | S |
| 87.15 | Preflight nije gate; runner nastavlja | Rizik | **Tačno** (Test_CoreTables... samo `LogFail` :199-201; runner ređa dalje :65-85) | P3 | `If m_Failed>0 Then Exit` posle preflight-a | S |
| 87.16 | Cross-zbirna audit preskače dangling/blank | Rizik | **Tačno** (:1291-1302 — poredi samo kad su OBA neprazna; `GetValueByKey`→Empty→skip) | P3 | Prijaviti dangling `OtpremnicaID` kao poseban count | S |
| 87.17 | „Expected to FAIL" komentar zastareo | Rizik | **Tačno** (header :21-22 vs. modSledljivost strict key sa `BrojZbirne`) | P3 | Obrisati komentar | S |
| 87.18 | Rezultat nije machine-actionable (Sub + MsgBox) | Rizik | **Tačno** (EndRun :1905-1927) | **P2** (koren 93.5) | `RunBusinessFlowProSuiteCore() As Boolean` (m_Failed=0) + wrapper | S |
| 87.19 | `Total` meša asercije/skip/fatal; scenario code zavisi od `m_Total` | Rizik | **Tačno** (:1936-1938, :1978-2013) | P3 | — | S |
| 87.20 | Test log = plain sheet sa Username; BuildGuard ga ne vidi | Rizik | **Tačno** (InitTestLog/AppendTestLog :2015-2047, `Environ$("Username")` :2046; plain range) | **P2** | Pokriveno 84.7 patch-om (skeniraj poznate log sheetove) | S |
| 87.21 | Soft-storno nije domen rollback; PASS i uz changed=0 | Rizik | **Tačno** (:1341-1377; LogPass bezuslovno :1372) | P3 | Preimenovati poruku + prikaz changed | S |
| 87.22 | Hard-delete ne briše fakture (FAK-/„N/god" bez TST-PRO); stavke se brišu → header bez stavki | Rizik | **Tačno** (marker za tblFakture samo `BrojFakture` :2282; GenerateBrojFakture modFaktura.bas:396 `N/god`; stavke preko `BrojPrijemnice` :2278) | **P2** | Pre brisanja stavki pokupiti njihove FakturaID → obrisati i headere | S |
| 87.23 | Ambalažni ledger ostaje (DokumentID=OTK-…) | Rizik | **Tačno** (TrackAmbalaza dobija `newID` = `OTK-` modOtkup.bas:535, :598-614; marker samo `DokumentID` :2302) | **P2** | Isti zahvat: registry stvarno kreiranih ID-jeva u run-u | M |
| 87.24 | Parcijalan hard-delete izgleda uspešno; substring marker | Rizik | **Tačno** (EH→0 + Debug.Print :2364-2367; `InStr` :2380; MsgBox total :2318) | P3 | Prikaz per-table grešaka u MsgBox | S |
| 87.25 | Cleanup ne uklanja seedove, mirror vozača, log sheet | Rizik | **Tačno** (HardDelete lista tabela :2278-2316 ne obuhvata master tabele; `ST-MIRTEST-90001` :1532) | P3 | Dokumentovati kao poznat ostatak | S |
| 87.26 | Dve `CreateSEFLive*` funkcije dupliraju tok; ne šalju na SEF | Rizik | **Tačno** (:2055-2256; nijedan SEF API poziv; `' ? ovde` :2170+) | P3 | Zadržati Dummy varijantu, obrisati drugu | S |
| 87.27 | Header ne navodi HardDelete/CreateSEFLive* | Rizik | **Tačno** (:24-32) | P3 | Dopuniti header | S |
| 87.28 | Šta je dobro rešeno (19 stavki) | Pozitivno | **Kontekst-Pozitivno** — potvrđeno (uklj. brisanje od dna, „BRISI" potvrda) | — | — | — |
| 87.29 | Ciljni dizajn (Host/Fixture/Runner/Cleanup registry) | Predlog | Premise tačne; L obim | — | SaveCopyAs-disposable model je pravi minimum | L |
| 87.30 | 33 prioriteta | Predlog | Tačne premise; minimalni paket = 3, 19, 24, 25 | — | — | — |
| 87.31 | Zaključak: jaka osnova, nebezbedna kao shipped makro/gate | Sinteza | **Tačno** | — | — | — |
| — | (87.7-dopuna) kontekstna beleška o „always-rollback TX" | — | **Netačno u kontekstu naloga, ne u FM** — FM je ovde tačniji od pre-registracije | — | Ispraviti internu belešku | — |

**Bilans FM-0081:** 31 podsekcija: 26 Tačno, 1 Dizajnersko ograničenje (87.10), 2 Kontekst-Pozitivno, 2 predloga/sinteze. Nema Netačno u FM-u; jedina korekcija ide na *pre-registrovani kontekst* (modul NEMA rollback TX). Eskalacije: **87.3/87.4 (environment guard u shipped test modulu — S)**, **87.18 (Boolean Core za gate — S)**, **87.20 (test log u artefaktu — pokriva 84.7 patch)**, **87.22/87.23 (hard-delete integritetske rupe — S/M)**.

### FM-0082 — `modCenovnik.bas` (137 linija; kod verifikovan u celini)

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 89.1 | Append-only dobar; temporalna semantika nedovršena | Sinteza | **Tačno** | — | — | — |
| 89.2 | Osnovni model dobar (dva modela cene razdvojena) | Pozitivno | **Kontekst-Pozitivno** (header :16-22) | — | — | — |
| 89.3 | `GetVazecaCena` nema `asOfDate` | P0 (FM 89.38) | **Tačno** (potpis :31-33); napomena: header :11 definiše „važeća = poslednji red", pa je implementacija verna *dokumentovanoj* nameri — gap je poslovni | P2 | `Optional asOfDate` + `Datum<=asOfDate`; forme prosleđuju datum dokumenta | M |
| 89.4 | Buduća cena važi odmah | P0 (FM) | **Tačno** (nema `Datum<=Date` filtera, :58-75) | P2 | Deo 89.3 patcha (`dv <= asOf`) | S* |
| 89.5 | Retroaktivni dokument dobija najnoviju cenu | P0 (FM) | **Tačno** (isti mehanizam) | P2 | Deo 89.3 patcha | S* |
| 89.6 | Naziv ne odgovara semantici | Rizik | **Tačno** | P3 | Rešava se 89.3 patch-om | — |
| 89.7 | Stale cena: `If c > 0` ne prazni polje → cena prethodnog proizvoda | P0 (FM) | **Tačno** (frmOtkup.frm:407-413 `If cI > 0 Then txtCena…`; frmDokumenta.frm:583-591 isto) | **P1** | Pre lookup-a obrisati auto-cena polja; na 0 upisati prazno + jasan hint | S |
| 89.8 | `0` objedinjuje sva failure stanja | Rizik | **Tačno** (:36, :83-85) | P3 | `TryGetVazecaCena(..., ByRef status)` uz zadržan stari potpis | M |
| 89.9 | `Datum` nije u schema guardu → tihi pad na „poslednji fizički red" | P0 (FM) | **Tačno** (:52 proverava cv/cS/ck/cc, ne `cD`; cD=0 → svi `dv=0` → `>=` bira poslednji) | **P2** | Dodati `Or cD = 0` u guard | S |
| 89.10 | Nevalidni datumi → dv=0, učestvuju u izboru | Rizik | **Tačno** (:63-67) | P3 | Deo 89.9/89.3 patcha (invalid = skip) | S |
| 89.11 | Isti datum → pobeda fizičkim redosledom (`>=`) | P0 (FM) | **Tačno** (:70); ublaženo: `AddCena` uvek appenduje na kraj → „poslednji unos pobeđuje" dok se tabela ne sortira ručno | P3 | Tie-break `CreatedAt` pa `CenaID` | S |
| 89.12 | Malformed najnoviji red poništava stariju validnu cenu | Rizik | **Tačno** (:77-79 IsNumeric samo na bestRow) | P3 | Uz typed status; ili preskočiti ne-numeričke kandidate | S |
| 89.13 | `AddCena` ne validira sortu/klasu/kulturu | Rizik | **Tačno** (:96-104 samo vrsta/klasa/cena>0) | P2 | Validirati klasa ∈ {I,II} + sorta obavezna | S |
| 89.14 | Prazna sorta bez wildcard semantike | Rizik | **Tačno** (exact-match :59-61) | P3 | Odlučiti: obavezna sorta (jednostavnije) | S |
| 89.15 | Nevalidan UI datum tiho postaje danas | P0 (FM) | **Tačno** (frmStammdaten.frm:2327 `If Not TryParseDateValue... Then datCen = Date`) | **P2** | MsgBox + fokus + prekid umesto fallback-a | S |
| 89.16 | Datum može nositi vreme; granularnost nedefinisana | Rizik | **Tačno** (CDbl(CDate) pun serial :66) | P3 | `DateValue()` u AddCena | S |
| 89.17 | Schema nije u startup self-heal-u | Rizik | **Tačno** (EnsureRuntimeSchema bez Cenovnika; EnsureCenovnikSchema ručno / EnsurePaletniListSchema modSetup.bas:959 / AdminEnsureEverything modAdmin.bas:261) | P3 | Dodati `EnsureCenovnikSchema` (bez MsgBox varijantu) u EnsureRuntimeSchema | S |
| 89.18 | Čitanje name-based, pisanje positional | Rizik | **Tačno** (:113-124 komentar „Redosled mora pratiti…") | **Prihvaćeno (AUD-003)** | Referenca na registrovani klaster | — |
| 89.19 | `AppendRow` neatomičan → parcijalan red | P0 (FM) | **Tačno** (modDataAccess: `ListRows.Add` pa per-cell; ErrHandler vraća 0, red ostaje) | **Prihvaćeno (AUD-003)** | Centralni fix u AppendRow (obriši newRow na EH) rešava i 91.26 | S |
| 89.20 | Nema conflict guard za isti ključ+datum | Rizik | **Tačno** | P3 | Upozorenje u frmStammdaten pre AddCena | S |
| 89.21 | Pogrešan red se ne može stornirati kroz UI | Rizik | **Tačno** (btnIzmeni sakriven frmStammdaten.frm:145; soft-delete dugme traži `Aktivan/Aktivna` — AktivanColName — a tblCenovnik ima `Stornirano`); ublaženo: novi red istog datuma appendovan kasnije pobeđuje (`>=`) | P3 | „Storniraj red" dugme za Cenovnik tab | M |
| 89.22 | Klik na istorijski red ne učitava datum | Rizik | **Tačno** (frmStammdaten.frm:1756-1766 učitava vrsta/sorta/klasa/cena; txtField4 ostaje/resetuje se na danas :2905) | P3 | Namerno? Dodati komentar ili učitati datum | S |
| 89.23 | PWA sync dokumentovan, ne implementiran | Rizik | **Tačno** (header modCenovnik.bas:24-26 „MORA"; StammdatenTabs = 13 tabova bez Cenovnika, modStammdatenSync.bas:19-35; nema `ExportCenovnik`) | P2 (uslovno — čim mobilni otkup koristi cenu) | `ExportCenovnik` po šablonu postojećih exportera | M |
| 89.24 | Nema verziju/atomsku publikaciju cenovnika | Rizik | **Tačno** | P3 | Posle 89.23 | M |
| 89.25 | Nema autorizacije na write boundary | Rizik | **Tačno** (AddCena bez provere; auth je opt-in faza) | P3 | `OblastAllowed(OBL_MATICNI)` provera u AddCena | S |
| 89.26 | Nema namenski monitoring event | Rizik | **Tačno** (samo LogErr :134) | P3 | `Monitor_Event "CENOVNIK_PRICE_ADDED"` | S |
| 89.27 | Dokument ne čuva provenance cene | Rizik | **Tačno** (forme dobijaju samo Double) | P3 | — | M |
| 89.28 | Ručni override tiho pregažen na Change | Rizik | **Tačno** (AutoFillCena pozvan sa 6/5 mesta na Change eventima) | P3 | Uz 89.7: prepisivati samo auto-popunjene vrednosti | M |
| 89.29 | Različita preciznost formi (`0.######` vs `0.00`) | Rizik | **Tačno** (frmOtkup.frm:408/413 vs frmDokumenta.frm:587-591) | P3 | Ujednačiti na `0.00` (ili zajednička konstanta) | S |
| 89.30 | VALIDACIJA_UNOSA=OFF dozvoljava cenu 0 u otpremnici | Rizik | **Dizajnersko ograničenje** — OFF je dokumentovani opt-out (frmDokumenta.frm:735-744; modConfig komentar :588-590) | P3 | — | — |
| 89.31 | `ExcludeStornirano` fail-open bez kolone | Rizik | **Tačno** (modHelpers: colStorno=0 → vrati data) | P3 | Cenovnik guard već pada na kolonama; dodati `Stornirano` u EnsureCenovnikSchema check | S |
| 89.32 | Nema plausibility limita | Rizik | **Tačno** (samo `cena > 0`) | P3 | Upozorenje na >X% promene | S |
| 89.33 | Performanse: pun scan po pozivu, bez indeksa | Rizik | **Tačno**; tabela mala | P3 | Ništa sada | — |
| 89.34 | Test coverage praktično nema | Rizik | **Tačno** (nijedan test modul ne poziva GetVazecaCena/AddCena) | P3 | 5-6 asercija u BusinessFlowPro | S |
| 89.35 | Dobre osobine (mali, storno-aware, centralni DataAccess…) | Pozitivno | **Kontekst-Pozitivno** — potvrđeno | — | — | — |
| 89.36 | Ciljni API (`ResolvePrice`/`PriceLookupResult`) | Predlog | Premise tačne; obim M/L | — | — | M–L |
| 89.37 | Minimalna bezbedna korekcija (12 koraka) | Predlog | **Tačno** — dobro pogođen minimal-delta; poklapa se s mojim predlozima | — | Usvojiti 1-9 kao jedan patch | M |
| 89.38 | Prioriteti P0/P1/P2 | Predlog | Premise tačne; FM-ov P0 je po ovoj kalibraciji P1 (89.7) + P2 (ostalo) — ništa nije aktivan gubitak podataka bez operaterske interakcije | — | — | — |
| 89.39 | Zaključak | Sinteza | **Tačno** | — | — | — |

**Bilans FM-0082:** 39 podsekcija: 31 Tačno, 1 Dizajnersko ograničenje (89.30), 2 Prihvaćeno-referenca (89.18, 89.19 → AUD-003), 2 Kontekst-Pozitivno, 3 predlozi/sinteze. Nema Netačno. Najjača pojedinačna eskalacija celog audita: **89.7 stale-price (P1, S-fix u obe forme)**; zatim P2 paket: `asOfDate` (89.3-5), `Datum` u guardu (89.9), UI datum fallback (89.15), validacija klase/sorte (89.13), PWA export (89.23 uslovno). FM-ov predlog 89.37 je praktično ispravan minimal-delta plan.

### FM-0083 — `modConfig.bas` — ponovljeni audit (974 linije; kod verifikovan u celini)

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 91.1 | Osnovni podaci (975 linija, blob, kritičnost) | Kontekst | **Tačno** (974+EOF; sadržaj odgovara) | — | — | — |
| 91.2 | Razlog ponovnog pregleda | Kontekst | **Kontekst-Pozitivno** (metodološki ispravno) | — | — | — |
| 91.3 | Četiri kategorije odgovornosti | Kontekst | **Tačno** | — | — | — |
| 91.4 | Tabela javnog API-ja (16 procedura) | Kontekst | **Tačno** — svih 16 potvrđeno u kodu (:766-973) | — | — | — |
| 91.5 | Šta je dobro rešeno | Pozitivno | **Kontekst-Pozitivno** — potvrđeno (fail-fast SetConfigValue :794-835 itd.) | — | — | — |
| 91.6 | **Tri config store-a; `GetConfigValue` hardkoduje literal `"tblSEFConfig"`** | Kritično | **Tačno** — modConfig.bas:770 `LookupValue("tblSEFConfig", "ConfigKey",…)` uprkos `TBL_SEF_CONFIG` (:58) i legacy `TBL_CONFIG` (:50); tri-store priča = pre-registrovan **AUD-018** | P3 (literal), šire Prihvaćeno (AUD-018) | Jednoslovna zamena literala konstantom u :770 | S |
| 91.7 | Startup validira legacy `tblConfig`, ne canonical store | Rizik | **Tačno** (modMain.bas ValidateAllTables niz sadrži `TBL_CONFIG`, nema `TBL_SEF_CONFIG`) | **P2** | Dodati `TBL_SEF_CONFIG` u niz (TBL_CONFIG ostaje — AUD-018 legacy-but-required) | S |
| 91.8 | Missing table/kolona/ključ/prazno = isti `""` | Rizik | **Tačno** (:766-778 + LookupValue vraća Empty za sve) | P3 | Typed read samo za kritične ključeve | M |
| 91.9 | Missing config menja poslovni režim bez greške | Rizik | **Dizajnersko ograničenje** — backward-compatible defaulti su dokumentovana odluka (:553-556, :847); rizik korektno opisan | P3 | Startup report korišćenih defaulta u SETUP health | M |
| 91.10 | `IsCloudSyncEnabled` permissive (typo→ON) | Rizik | **Tačno** (:853-860 `Case Else` → True) | P3 | Strict parse + LogWarn za nepoznatu vrednost | S |
| 91.11 | `ConfigFlag` skriva nevalidne vrednosti | Rizik | **Tačno** (:886-897 `Case Else` → defaultOn) | P3 | LogWarn na nepoznatu ne-praznu vrednost | S |
| 91.12 | Duplikati ključeva nekontrolisani; prvi red pobeđuje | Rizik | **Tačno** (LookupValue prvi match; SetConfigValue :809-815 prvi match) | P3 | Duplicate-key provera u RunSetupHealthCheck | S |
| 91.13 | Ključevi praktično case-sensitive | Rizik | **Tačno** (CStr poređenja; nigde `Option Compare Text` — grep prazan) | P3 | UCase$ kanonizacija u Get/Set | S |
| 91.14 | `SetConfigValue` nije domenski validator | Rizik | **Tačno** (:794-829) | P3 | Validaciju držati u UI editoru (postojeće) | — |
| 91.15 | Nema centralni metadata registry ključeva | Rizik | **Tačno** (modPodesavanja `CfgAdd` literali — :57, :120, :133) | P3 | ConfigEditorFields JE de-facto registry; dopunjavati njega | — |
| 91.16 | `tblSEFConfig` pogrešno ime za centralni store | Rizik | **Tačno** (sadrži Google/monitoring/licencu — modSetup:593, :807-823) | Prihvaćeno (AUD-018; rename skup) | Ne dirati fizičko ime | — |
| 91.17 | Secrets plaintext u workbooku | Rizik | **Dizajnersko ograničenje** Excel/VBA platforme (VeryHidden modPodesavanja.bas:719 je hygiene) | P3 | — | — |
| 91.18 | Hardkodovana publish šifra nije autentikacija | Rizik | **Tačno**, ali kod to i tvrdi („dev gate… sprecava slucajnu objavu", modConfig.bas:19-21; poređenje InputBox modAdmin.bas:280-284); komentar „PROMENI pre isporuke" nije ispoštovan | P3 | Ili promeniti vrednost po instrukciji, ili izbrisati lažno obećanje iz komentara | S |
| 91.19 | Folder ID-jevi compile-time; `BACKUP_FOLDER_ID` bez VBA potrošača | Rizik | **Tačno** (grep: samo gas/DriveFolder.gs referencira *property name*, ne konstantu) | P3 | Komentar „rezervisano za GAS backup" ili ukloniti | S |
| 91.20 | `APP_VERSION` i build identitet nisu jedan ugovor | Rizik | **Tačno** (:13 vs modBuildInfo placeholder) | P3 | Pokriveno publish-guard patch-om (86.9) | S |
| 91.21 | `G:\My Drive` developerski fallback | Rizik | **Tačno** (:517-519; fallback u modBankaImport.bas:1041-1049) | P3 | Prazan default + poruka „pokreni SetupBankFolders" | S |
| 91.22 | pdftotext bez platform policy | Rizik | **Dizajnersko ograničenje** (Windows-only aplikacija) | P3 | — | — |
| 91.23 | Error-code collision `vbObjectError+2700` | Rizik | **Tačno** — modConfig.bas:764 (`ERR_STORNO_FW_BASE`, komentar „ne preklapa se") vs modBankaImport.bas:46 (`ERR_BIM_IMPORT_BASE`, aktivno korišćen :207-224); detalj „banka konstante deluju neiskorišćeno" je netačan — koriste se | P3 | Pomeriti banka bazu (npr. +2900 posle provere) + ispraviti komentar u modConfig | S |
| 91.24 | Dormantne konstante (PROIZVODJACI/HLADNJACA/LAGER/KVALITET/SLEDLJIVOST) | Rizik | **Tačno** (grep: svih 5 samo u modConfig.bas) | P3 | Komentar `' RESERVED (bez schema/CRUD)` | S |
| 91.25 | Centralizacija nepotpuna | Rizik | **Tačno** | P3 | — | — |
| 91.26 | Config upsert nije atomican (insert put) | Rizik | **Tačno** (AppendRow partial-row; SetConfigValue :824-828 raise posle) | P3 (fix zajednički sa AUD-003) | Centralni AppendRow EH-cleanup | S |
| 91.27 | Nema config change journal | Rizik | **Delimično** — AppendRow poziva `WriteJournalRow` + `StampRowAudit` (generic trag na insertu); update put (`RequireUpdateCell`) nema old→new zapis, što je jezgro tvrdnje | P3 | — | M |
| 91.28 | Javni setter bez authz | Rizik | **Dizajnersko ograničenje** VBA (svaki modul ionako može pisati u sheet) | P3 | — | — |
| 91.29 | VeryHidden je hygiene, ne anti-tamper | Rizik | **Tačno** (modPodesavanja.bas:719) | P3 | Ažurirati komentar koji ga zove anti-tamper | S |
| 91.30 | `APP_NAME` „OtkupApp" vs AgriX branding | Rizik | **Tačno** (:12) | P3 | Odluka vlasnika; ne dirati kod | — |
| 91.31 | Performance: pun read po pozivu | Rizik | **Tačno**; tabela mala, GetTableData ima request-scoped keš | P3 | — | — |
| 91.32 | Failure scenariji A-E | Sinteza | **Tačno** — svi izvedeni iz potvrđenih mehanizama | — | — | — |
| 91.33 | Šta NE raditi (bez masovnog refaktora) | Pozitivno | **Kontekst-Pozitivno** — ispravna kalibracija | — | — | — |
| 91.34 | Ciljni ugovor (ConfigDefinition/ReadResult/WriteResult) | Predlog | Premise tačne; obim L za VBA | — | — | L |
| 91.35 | Prioriteti P0/P1/P2 | Predlog | Premise tačne; FM-ov „P0" je po ovoj kalibraciji P2 (91.7) + P3 — nijedna stavka nije aktivan kvar | — | Minimalni paket: 91.7 + 91.6-literal + 91.23-komentar | S |
| 91.36 | Korekcija prvog zaključka (FM-0001) | Sinteza | **Tačno** — fer revizija | — | — | — |
| 91.37 | Konačni profil | Sinteza | **Tačno** | — | — | — |

**Bilans FM-0083:** 37 podsekcija: 24 Tačno (od toga 91.23 sa netačnim sporednim detaljem o „neiskorišćenim" banka konstantama), 1 Delimično (91.27 — journal na insertu postoji), 4 Dizajnersko ograničenje (91.9, 91.17, 91.22, 91.28), 1 Prihvaćeno (91.16 → AUD-018), 4 Kontekst-Pozitivno/Kontekst, 3 predlozi/sinteze. Nema Netačno kao celina. Prior audit („sound") ostaje validan za katalog konstanti; jedina prava eskalacija je **91.7 (dodati `TBL_SEF_CONFIG` u ValidateAllTables — S)** + kozmetički S-fix literala (91.6).

### FM-0084 — `modE2EReleaseGate.bas` (157 linija; kod verifikovan u celini)

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 93.1 | Osnovni podaci | Kontekst | **Tačno** | — | — | — |
| 93.2 | Namera: orkestrator 6 VBA suite-ova + 3 warn | Kontekst | **Tačno** (:28-53) | — | — | — |
| 93.3 | Javni API bez rezultata/raise/receipt | Kontekst | **Tačno** (:23-65) | — | — | — |
| 93.4 | Šta je dobro rešeno | Pozitivno | **Kontekst-Pozitivno** — potvrđeno | — | — | — |
| 93.5 | **Normalan povratak makroa = automatski PASS** | Kritično | **Tačno** — modE2EReleaseGate.bas:74-77: `Application.Run procName` pa bezuslovno `E2E_Pass … "Verify suite summary output/log…"`; jedini FAIL uslov je neuhvaćena VBA greška (:80-82) | **P2** | Ili deprecirati modul (nije ni u proceduri — 93.13), ili `*Core() As Boolean` po suite-u + `E2E_Pass` samo na True | S–M |
| 93.6 | Svi pozvani suite-ovi interno gutaju failure | Rizik | **Tačno** — potvrđeno za svih 6: EndGoogleSmokeRun (MsgBox, bez raise), RunMasterSyncSmokeSuite (isti finisher + tx rollback), FinishNovacSuite (modNovacTests:216-234), FinishFakturaSuite (modFakturaTests:585+), EndRun (BusinessFlowPro :1905-1927), EndHealthRun (modProductionHealthCheck:1186+) | P2 (isti paket) | `m_Failed=0` izlaz iz svakog finishera | M |
| 93.7 | Lažno-zeleni scenario TOTAL=9/PASS=6/FAIL=0 uz 6 crvenih suite-ova | Rizik | **Tačno** (direktna posledica 93.5+93.6 — ne eskalirati zasebno) | — | — | — |
| 93.8 | `m_Total` nije broj testova | Rizik | **Tačno** (:135-155 broji korake) | P3 | — | — |
| 93.9 | Čist PASS nedostižan (WARN≥3 bezuslovno) | Rizik | **Tačno** (:46-53 dva ManualGate + jedan Warn; PASS grana :128-130 samo uz m_Warn=0) | P3 | Warn-ove pretvoriti u checklistu u poruci | S |
| 93.10 | `E2E_ManualGate` = alias za Warn; nema dokaza izvršenja | Rizik | **Tačno** (:84-86) | P3 | InputBox potvrda „PASS/FAIL" po GAS koraku | S |
| 93.11 | Hardkodovani „29 handlers" / „10/10" | Rizik | **Tačno** (:47, :50) | P3 | Tekst bez brojeva ili čitati iz GAS statusa | S |
| 93.12 | Unapred dodat ProductionHealth waiver je blanket | Rizik | **Tačno** (:52-53 bezuslovno) | P3 | Vezati za check ID ili ukloniti | S |
| 93.13 | Gate nije povezan sa objavom | Rizik | **Tačno** (jedina pominjanja van modula: docs changelog; `PublishReleaseToDrive` i release.sh ga ne zovu) | P3 (ali relativizuje ceo modul) | Ubaciti u release.sh checklist tekst ILI deprecirati | S |
| 93.14 | Ni registrovan FAIL ne blokira (Sub, bez raise) | Rizik | **Tačno** (:120-131; EH :58-64 bez re-raise) | P2 (isti paket kao 93.5) | `Err.Raise` posle EndE2EGate kad m_Fail>0 | S |
| 93.15 | `Application.Run` stringovi bez compile-time veze | Rizik | **Tačno** (:75) | P3 | Direktni pozivi posle RF-01 čišćenja | S |
| 93.16 | Nejedinstvena imena: `modNovacTest(+s)`, `modFakturaTest(+s)` sa istim Public Sub | Rizik | **Tačno** — sva 4 fajla postoje na a0bc9e2, oba para definišu `RunNovacSmokeSuite`/`RunFakturaSmokeSuite` (linija 16 u svakom); `Application.Run` nekvalifikovano → nedeterminizam | **Prihvaćeno (AUD-016; duplikati obrisani u RF-01)** | — | — |
| 93.17 | Suite je mutaciona (Google spreadsheet, Novac/Faktura redovi, `APP_LAST_HEALTHCHECK_AT`) | Rizik | **Tačno** (CreateSpreadsheet `TST-GOOGLE-SMOKE-*` + TrashGoogleDriveFile; modFakturaTests CreateFaktura_TX/SaveNovac; modProductionHealthCheck.bas:1202) | P3 | — | — |
| 93.18 | Pokretanje na build-masteru kvari blanko artefakt (TOCTOU sa BuildGuard) | Rizik | **Tačno** (log sheetovi plain-range — nevidljivi guardu; TST redovi vidljivi) | P3 | Napomena u RELEASE_PROCEDURE: E2E samo na disposable kopiji | S |
| 93.19 | Nema environment guard | Rizik | **Tačno** | P3 | Deli guard iz 87.3 | S |
| 93.20 | Redosled: health poslednji, posle mutacija | Rizik | **Tačno** (:28-44) | P3 | Health pre i posle | S |
| 93.21 | Nema dependency-aware SKIP/BLOCKED | Rizik | **Tačno** | P3 | — | M |
| 93.22 | Modalni child MsgBox-ovi onemogućavaju automatizaciju | Rizik | **Tačno** (svi finisheri MsgBox) | P3 | `showUi` parametar u Core varijantama | M |
| 93.23 | Nema build/artifact identitet u rezultatu | Rizik | **Tačno** (nigde APP_VERSION/BUILD_* u modulu) | P3 | Dodati u Begin/End ispis | S |
| 93.24 | Nema zajednički RunID/receipt | Rizik | **Tačno** (modul nema m_RunID; child suite-ovi imaju sopstvene) | P3 | — | M |
| 93.25 | Pokrivenost = istorijski v6.10 snapshot | Rizik | **Tačno** (naslov v6.10 :8, :99 vs `APP_VERSION="2.21.0"` modConfig.bas:13 i RELEASE_GATES.md:14 „v6.23 production handoff") | P3 | Deprecirati ili uskladiti sa RELEASE_GATES.md | S |
| 93.26 | Nema changed-surface politiku | Rizik | **Tačno** (RELEASE_GATES.md bira po površinama; modul fiksna lista) | P3 | — | — |
| 93.27 | Error-reporting put nezaštićen | Rizik | **Tačno** (:134-156 LogInfo/LogWarn/LogError bez lokalnog handlera) | P3 | `On Error Resume Next` oko log poziva | S |
| 93.28 | Ciljna arhitektura (contract/policy/enforcement) | Predlog | Premise tačne; L obim | — | — | L |
| 93.29 | Minimalni popravni paket (14 koraka) | Predlog | **Tačno** — koraci 2-5 su pravi minimal-delta; ostalo opciono | — | Usvojiti 2-5 + 6-7 (RF-01 već pokriva 7) | M |
| 93.30 | Meta-testovi gate-a | Predlog | Premise tačne; preskočiti dok se ne odluči sudbina modula | — | — | M |
| 93.31 | Prioriteti P0/P1/P2 | Predlog | Premise tačne; FM-ov „P0" je ovde P2 jer gate nije u aktivnoj proceduri (93.13) | — | Prvo odluka: deprecate vs. popravka | — |
| 93.32 | Šta ne treba raditi | Pozitivno | **Kontekst-Pozitivno** — ispravna kalibracija | — | — | — |
| 93.33 | Konačna procena: dashboard, ne gate | Sinteza | **Tačno** | — | — | — |

**Bilans FM-0084:** 33 podsekcije: 24 Tačno, 1 Prihvaćeno (93.16 → AUD-016/RF-01), 2 Kontekst-Pozitivno, 6 kontekst/predlozi/sinteze. Nema Netačno, nema Delimično. Kritični nalaz 93.5 je precizno potvrđen (modE2EReleaseGate.bas:74-77) i realan, ali je hitnost P2 a ne viša jer modul **nigde nije pozvan** u aktivnoj release proceduri (93.13) — prva odluka treba da bude deprecate-ili-popravi; ako se zadrži, minimum je Boolean `Core` ugovor po suite-u (veže se na 87.18) + `Err.Raise` na FAIL.

---

## Ukupni zaključak audita (svih 187 podsekcija verifikovano protiv koda)

- **Tačnost dokumenta je vrlo visoka:** ~142 Tačno, 2 Delimično (84.3 Alt+F8 nijansa; 91.27 journal na insertu postoji), 0 Netačno kao celina (2 netačna *sporedna detalja*: 91.23 „banka konstante neiskorišćene"; kontekstna beleška o TX-rollbacku za FM-0081 važi za MasterSync, ne za ovaj modul), 8 Dizajnersko ograničenje, 4 Prihvaćeno (AUD-003 ×2, AUD-016, AUD-018), ostalo Kontekst-Pozitivno/sinteze/predlozi.
- **Jedina P1 eskalacija:** 89.7 — stale cena u `frmOtkup`/`frmDokumenta` (`If c > 0` ne prazni polje; frmOtkup.frm:407-413, frmDokumenta.frm:583-591). S-fix.
- **P2 paketi (mali delta, veliki efekat):** (1) publish-guard u `PublishReleaseToDrive` — placeholder/dirty deny + disk↔workbook SHA cross-check (84.12, 86.3, 86.8, 86.13); (2) BuildGuard scan poznatih plain-range logova — `SETUP_LOG`, test logovi (84.7, 87.20); (3) environment guard + Boolean `Core` u test suite-ovima (87.3, 87.4, 87.18, 93.5/93.6/93.14); (4) hard-delete rupe za fakture/ambalažu (87.22, 87.23); (5) Cenovnik `asOfDate` + `Datum` guard + UI datum fallback (89.3-5, 89.9, 89.15); (6) `TBL_SEF_CONFIG` u `ValidateAllTables` (91.7).
- Sve „Kritično"-naslovljene podsekcije (84.7, 91.6, 93.5) potvrđene su uz citate fajl:linija u tabelama.

---


---
---

# DEO III — Delta trijaža Functional Map v142 (FM-0085…FM-0140)

**Datum:** 2026-07-20
**Izvor:** `AgriX_Functional_Map` v142 — novi/nepokriveni unosi u opsegu FM-0085…FM-0140.
**Metod:** identičan DEO I/II (svaka stavka pojedinačno protiv koda; Kritičan/Visok/„P0" uz citat fajl:linija; 8 paralelnih prolaza, blokovi K1…K8).
**Sidro koda:** `origin/main` **v2.24.0** (`9fd7087`), provereno u zasebnoj worktree kopiji. Header v142 tvrdi sidro `a0bc9e2`, ali fajlovi koje popisuje (`modStornoWarm`, `modTestStornoCentar`, prošireni `modStornoFlow`) postoje tek od v2.24.0 — zato je cela delta verifikovana protiv `origin/main`, ne protiv navedenog sidra.

**Obuhvat delte:** 38 FM unosa (novi fajlovi ili fajlovi neanalizirani u v35/v85). **Preskočeno kao već-analizirano ili duplo:** FM-0085, FM-0086 (frmOtkup/frmDokumenta — DEO II 89.x), FM-0094, FM-0095 (dupli test opisi), FM-0109 frmLogin, FM-0110 modAuth, FM-0115 frmSplash, FM-0117…FM-0120 clsStmBtn/clsConfigBtn/clsAdminBtn/clsLookupMenuBtn, FM-0122 clsBlokIsplata, FM-0124…FM-0126 SEF DTO klase, FM-0135 modKvalitet (stub), FM-0138 modML (stub), FM-0139 modFakturaTests — svi pokriveni ranije ili bez novog sadržaja.

## ⚠️ Napomena o E2E gate-u (usklađivanje sa DEO II / AUD-039)

Više test-modula u ovoj delti (FM-0093/0097/0098/0099/0100/0136) FM markira „P0 E2E false-green" jer suite završava normalno i kad interno padne. Verifikovano: `modE2EReleaseGate.bas:23-49` **zaista poziva** `RunGoogleSyncSmokeSuite`/`RunMasterSyncSmokeSuite`/`RunNovacSmokeSuite`/`RunFakturaSmokeSuite`/`RunBusinessFlowProSuite`/`RunProductionHealthCheck` — pa je lažni-zeleni lanac **realan po mehanizmu**. **ALI:** `modE2EReleaseGate`/`E2E_RunVbaSuite` **nije pozvan nigde** u aktivnoj release proceduri (grep prazan; `PublishReleaseToDrive` ne zove gate). Zato je ceo lanac **latentan = AUD-039**, kalibrisan na **P1**, ne P0 (isto kao DEO II FM-0084/93.13). Minimalni popravak ostaje isti: suite-ovi dobijaju Boolean `Core`/`Err.Raise`, gate čita rezultat umesto completion-a.

## Zbirni bilans delte (38 fajlova, 8 blokova)

Dominantno **Tačno** (v142 je i dalje činjenično vrlo precizan). Glavna korekcija je, kao i ranije, kalibracija težine: sistematski „P0/P1" oznake padaju na **P2/P3** zbog (a) single-writer desktop modela (read-only dijagnostike, fail-open readovi, whole-table test rollback = Prihvaćeno/P2), (b) 4-nivo integriteta banka-parsera (drift = prekid uvoza, ne tiha korupcija), (c) kozmetičkih slojeva (theme, mouse-wheel) bez data-rizika, (d) test-modula koji nisu u aktivnom E2E gate-u. Veliki deo delte mapira na **već registrovane** nalaze (AUD-002, AUD-003, AUD-007, AUD-016, AUD-017, AUD-018, AUD-034, AUD-037, AUD-039).

**Novi čist P0 (aktivan gubitak/korupcija):** **0.** Najozbiljniji potvrđeni lanac (FM-0093 E2E false-green) je latentan (AUD-039); ostali „P0" su ili već registrovani (AUD-002) ili single-writer-Prihvaćeno.

**Novi P1 klasteri (verifikovani, minimal-delta popravke):**
- **Agrohemija — cena (najvredniji nalaz delte):** `frmAgrohemija` izlaz snapshot-uje cenu u korpi ali je NE prosleđuje `SaveMagacin` kao `overrideCena` (`frmAgrohemija.frm:623-630`) → `SaveMagacin` ponovo čita master cenu; **ulaz ISPRAVNO prosleđuje** (`:843`) — asimetrija dokazuje oversight. Fix S (jedan argument). `modAgrohemija` uz to tiho upisuje `Cena=0/Vrednost=0` kad je cena nenumerička (`:109-130`) → potcenjen dug. Fix S.
- **Brojevi/sync — duplikati:** `modBrojevi.GenerateBrojPrijemnice` EH vraća validan-looking `1/ddmmyy` duplikat (`:203`) umesto hard-fail-a (fix S); `modMasterSync.GenerateBrojZbirne` (`:2887-2928`) je paralelni **row-count** generator (`seq=count+1`) koji na rupama pravi duplikat (`1/ddmmyy` + `…-3` → ponovo `-3`) umesto canonical `SuggestNextBroj`/`MaxSeqFromTable` (fix M delegiranjem).
- **MasterSync — pogrešan upis/writeback:** `TryUpdateVozacID` vraća True i na neuspeh write-a (`:1773-1798`) → GS `Synced>Master`, `VozacID` prazan (fix S); nevalidan datum → **današnji datum** na oba puta (OTK `:1547-1550`, VOZ `:2598-2600`, fix S); auto-otpremnica grupiše po `Stanica|Datum|Vozac|Klasa` pa meša vrste/cene/ambalažu u 1 otpremnicu (`:668-672`, fix M); VOZ link (`LinkZbirnaToOtkupAndOtpremnica :2701-2771`) ne proverava membership i prepisuje postojeće veze bez konflikt-politike (fix M); poison spreadsheet na neuspeh header write-a (`:476-494`, fix M).
- **Integritet — lažni zeleni:** `modIntegritet.WriteErr` (`:1304-1310`) ne diže `m_totalIssues`, a overlay/MsgBox (`:59,:90`) čitaju samo taj brojač → moguće „0 neusklađenih" uz GRESKA blokove; `Empty` = i PASS i ERROR (`:84-85`). Fix M (uvesti ErrorCount + typed rezultat).
- **Sledljivost — nepotpun trag:** `TraceByZbirna` filtrira po pomoćnom `OtpremnicaID` umesto canonical `tblOtkup.BrojZbirne` (`modSledljivost.bas:540-544`), a normalizacija broja nekonzistentna (`:282` vs `:464`) → moguć nepotpun trag; `frmSledljivost`/PDF prikazuju nepotpun trace kao kompletan (P1-a/f). Upis ipak ide kroz proveren TX (nema korupcije). Fix S–M.
- **Stanica-mirror:** missing-shadow (stanica bez `tblVozaci` para) → `modMasterSync.StampVozacFromStanicaForMalina` i `modAutoHladnjaca` bezuslovno `vozacID=stanicaID` → dokument dobija FK bez `tblVozaci` reda (`modMalina` 116.6/116.20/116.21). Fix: jedan canonical `IsManagedStationMirror` + re-raise u Ensure EH (M).

**Novi P2 paketi (mali delta, veliki efekat):** `modProductionHealthCheck` SEF lista koristi nepostojeći `SEF_CANCELLED` i propušta `SEF_REJECTED/SYNC_ERROR/TECH_FAILED` (`:871`, drift vs `modConfig.bas:659-663`) + ponovljeni „parent OK posle child FAIL" (`:951`, `:928`); `modParse` single-separator factor-1000 rizik (121.6, = AUD-007 familija); shipped destruktivni test runneri bez environment guard-a (`modTestStornoCentar` 120.9 — nula potvrde nad 23 `Public` makroa; AUD-039 familija); `modStornoWarm` lažni scheduled/cancelled state (`:51-54`, `:118-124`).

**Sistematski precenjeno (FM „P0/P1" → realno P2/P3):** `modMouseWheel` lifecycle (kozmetička off-by-default funkcija), `modTheme` ceo sloj (samo boje, nema data-loss — DisableField/DisableCombo clear je namerni mode-switch), `frmMarza` (legacy/nekorišćena forma), permission-bypass u UI adapterima (AUTH opt-in/OFF + soft model + nezaštićen `.xlsm`), banka-parseri (4-nivo integritet hvata drift).

**Format napomena:** blokovi K1…K8 slede; svaka FM stavka je zaseban red (narativ → numerisane podsekcije NN.x; risk-tabele → red po redu), sa kolonama Opravdanost / Hitnost / Predlog / Napor identičnim DEO I/II.

---

## v142 blok K1 — modProductionHealthCheck, modSchemaGuard, modIntegritet (FM-0087/0089/0090) [sidro origin/main v2.24.0]

I have completed verification against the code. All three FM entries (FM-0087, FM-0089, FM-0090) are narrative (numbered subsections, no risk tables). Below is the full per-item audit.

Key code confirmations: SEF constants in `modConfig.bas:653-663` (there is **no** `SEF_CANCELLED`; `SEF_REJECTED`/`SEF_SYNC_ERROR`/`SEF_TECH_FAILED`/`SEF_READY`/`SEF_SENDING`/`SEF_UNKNOWN` all exist and are uncovered); `GetSpreadsheetID` returns `""` on HTTP≠200/no-token/exception (`modGoogleSheets.bas:1382,1406,EH`); `GetTable` compares names case-sensitively (`modDataAccess.bas:79`); error codes 7300/7400 are unique to modSchemaGuard; no `modIntegritetTests`/`modSchemaGuardTests` exist.

---

### FM-0087 — `modProductionHealthCheck.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 99.1 | Osnovni podaci (13 check grupa, entry point) | Osn. podaci | Kontekst-Pozitivno | — | Kontekst — nije rizik | — |
| 99.2 | Redosled 13 provera | narativ | Kontekst-Pozitivno | — | Kontekst — nije rizik | — |
| 99.3 | Glavni zaključak (nije release gate) | narativ | Kontekst | — | Kontekst — nije rizik | — |
| 99.4 | Šta je dobro rešeno | Pozitivno | Kontekst-Pozitivno | — | Kontekst — nije rizik | — |
| 99.5 | „read-only" nije potpuno tačno | narativ | **Tačno** (header:12 vs write :1202,:1247) | P3 | Precizirati header; side-efekti = AUD-006/AUD-018 | S |
| 99.6 | Rezultat nije machine-readable ni blocking | narativ | **Tačno** (:28 Sub, nema Bool/Err na m_Fail) | P2 | Dodati `...Core() As Result`; UI wrapper prikazuje | L |
| 99.7 | Kontradiktoran Google parent OK posle child FAIL | **Kritičan** | **Tačno** (:951 bezuslovni `HealthOk` posle :945-949) | P2 | :951 uslovi parent OK na child delta-brojače | S |
| 99.8 | Folder probe false-green na HTTP/auth grešku | narativ | **Tačno** (:1024-1031; `GetSpreadsheetID` ""→OK, modGoogleSheets.bas:1406) | P2 | Vratiti typed DriveLookupResult; OK samo na NOT_FOUND | M |
| 99.9 | Google se proverava i kad je sync OFF | narativ | **Tačno** (config/auth pre flag :945-949) | P2 | Prvo proveriti flag; N/A kad sync off | M |
| 99.10 | APP_SETUP_COMPLETED = presence-only | narativ | **Tačno** (:1096 Len=0 else OK) | P2 | :1096 typed Boolean policy umesto neprazno | S |
| 99.11 | Auth check nije pasivan (GetAccessToken refresh) | narativ | **Tačno** (:992) | P3 | Odvojiti stored-token od active-probe; = AUD-006/018 | M |
| 99.12 | LAST_HEALTHCHECK_AT piše i za FAIL run | narativ | **Tačno** (:1202 bezuslovno) | P2 | Dodati STATUS/PASS_AT ključeve pored AT | S |
| 99.13 | Config write kreira red + journal | narativ | **Tačno** (:1202→AppendRow) | P3 | Već pokriveno AUD-003/AUD-006/AUD-018 | S |
| 99.14 | Schema check javlja samo prvi nedostatak | narativ | **Tačno** (:62-108 raise na prvom) | P2 | CollectMissing prolaz za grupni FAIL | M |
| 99.15 | Core inventory nepotpun za ostatak modula | narativ | **Tačno** (:65-72 8 tabela; kasnije BankaImport/Parcele) | P3 | Deklarativni schema manifest po domenu | M |
| 99.16 | Duplicate coverage bez StavkaID/business brojeva | narativ | **Tačno** (:116-123) | P2 | Dodati StavkaID, BrojFakture/Prijemnice dup | M |
| 99.17 | Prazni PK preskočeni bez zasebnog checka | narativ | **Tačno** (:1509 GoTo NextRow; nema blank-PK) | P2 | Dodati `Check_RequiredPrimaryKeysNotBlank` | M |
| 99.18 | Duplicate dict bez CompareMode | narativ | **Tačno** (:1498 binary) | P3 | vbTextCompare uz log originalnih varijanti | S |
| 99.19 | ActiveRowExists/IsStorniranoRow first-match | narativ | **Tačno** (:1379-1391) | P2 | Vratiti typed multi-status pod duplikatima | M |
| 99.20 | HealthNumeric tiho pretvara u nulu | narativ | **Tačno** (:1346-1352) | P2 | Vratiti (IsValid,Value) ili odmah FAIL na non-num | M |
| 99.21 | Novac integrity preuzak | narativ | **Tačno** (:177-198 samo 4 pravila) | P2 | Dodati domain/FK/dup NovacID provere | M |
| 99.22 | Payment sum helperi gutaju greške→0 | narativ | **Tačno** (:1595,:1628 EH→0) | P2 | Razlikovati infrastrukturnu grešku od 0 uplata | M |
| 99.23 | FakturaStavke ref = samo existence | narativ | **Tačno** (:247-257) | P2 | Dodati dup StavkaID, zbir-vs-Iznos, backlink | M |
| 99.24 | Prijemnica/Faktura flags jednosmerni | narativ | **Tačno** (:309-319) | P2 | Bidirekcioni graph audit | M |
| 99.25 | Faktura payment provera samo jedne strane | narativ | **Tačno** (:380-394 samo Placeno grana) | P2 | Simetrična expected-status rekonstrukcija | M |
| 99.26 | Otkup payment isti jednosmerni problem | narativ | **Tačno** (:457-471) | P2 | Isto: expected status/date iz uplata | M |
| 99.27 | Kooperant reconciliation vredna | Pozitivno | Kontekst-Pozitivno | — | Kontekst — nije rizik | — |
| 99.28 | Reconciliation koristi iznos kao uslov (offset) | narativ | **Tačno** (:589 `unattributed<=0.005`) | P2 | :589 uslov na count, iznos = impact metrika | S |
| 99.29 | Reconciliation bez konkretnih redova | narativ | **Tačno** (:590-599 agregati) | P3 | Logovati prvih 25 OtkupID kao duplicate check | S |
| 99.30 | Iznos recon koristi live Prijemnicu, ne frozen Stavke | narativ | **Tačno** (:641-666 live sum) | P2 | 3-nivo audit uklj. tblFakturaStavke snapshot | M |
| 99.31 | Financial drift samo WARN | narativ | **Tačno** (:701,:713 HealthWarn) | P2 | Severity po stanju (finalized/SEF→FAIL) | S |
| 99.32 | Cross-zbirna preskače blank-vs-populated | narativ | **Tačno** (:785 oba Len>0) | P2 | Blank na jednoj strani = mismatch | S |
| 99.33 | GetValueByKeySafe skriva grešku kao blank | narativ | **Tačno** (:1452-1453 EH→Empty) | P3 | Vratiti status uz vrednost | S |
| 99.34 | Cross-zbirna ne pokriva ceo invariant | narativ | **Tačno** (:769-799 samo BrojZbirne) | P3 | Dodati reverse/klasa/stanica konzistentnost | M |
| 99.35 | SEF lista hardkodovana, driftovala | narativ | **Tačno** (:871 `SEF_CANCELLED` nepostojeći; modConfig.bas:659-663 REJECTED/SYNC_ERROR nepokriveni) | P2 | :871 koristiti WF_SEF_* konstante + state matrica | M |
| 99.36 | SEF optional kolone nestaju, audit prođe | narativ | **Tačno** (:838-849 GetColumnIndex skip) | P2 | Required kolone za canonical outbound schema | S |
| 99.37 | SEF samo presence, ne konzistentnost | narativ | **Tačno** (:870-889) | P3 | Dodati dup docID/format/status matcheve | M |
| 99.38 | Stornirane Fakture potpuno preskočene u SEF | narativ | **Tačno** (:861 GoTo NextRow) | P2 | Poseban invariant za stornirane sa remote docID | M |
| 99.39 | Soft-delete coverage parcijalan (3 veze) | narativ | **Tačno** (:913-926) | P2 | Deklarativna matrica veza (+Novac,FakturaStavke) | M |
| 99.40 | CountActiveRef ne propagira failure→parent OK | narativ | **Tačno** (:928 OK iako helper :1161 loguje FAIL) | P2 | Helper vraća status; parent gate na njega | S |
| 99.41 | „MasterSchema" = lokalne tabele, ne remote | narativ | **Tačno** (:1039-1072 lokalno) | P3 | Preimenovati + dodati remote schema probe | M |
| 99.42 | Google config vremenski kontradiktoran | narativ | **Tačno** (:970-971 pre refresh :992) | P3 | Proveriti metadata posle refresh-a | S |
| 99.43 | Health log van blank-build guard-a | narativ | **Delimično** (log = range :1247-1252 potvrđeno; AssertBlankBuild zavisnost) | P2 | Uključiti log-sheetove u build guard ili čistiti | S |
| 99.44 | Log write failure nevidljiv | narativ | **Tačno** (:1242,:1259 On Error Resume Next) | P2 | Uvesti LogStatus u rezultat | S |
| 99.45 | RunID nije vezan za artefakt | narativ | **Tačno** (:1169) | P3 | Dodati BUILD_SHA, APP_VERSION, workbook hash | S |
| 99.46 | Rnd bez Randomize | narativ | **Tačno** (:1169) | P3 | Centralni GUID/ID helper | S |
| 99.47 | Summary Total = findings, ne checkovi | narativ | **Tačno** (:1218,:1226,:1234) | P3 | Razdvojiti CheckCount/FindingCount | S |
| 99.48 | Prazne tabele → mnogo WARN | narativ | **Tačno** (višestruki IsEmpty→WARN) | P3 | Uvesti profile (BLANK/NEW_CLIENT/ACTIVE) | M |
| 99.49 | Performance O(N×M) | narativ | **Tačno** (ActiveRowExists re-čita target; nema BeginTableCache) | P2 | Jednom izgraditi ID/ref/payment indekse | M |
| 99.50 | Audit nije platform-neutralan (WinHttp) | narativ | **Delimično** (zavisnost modGoogleSheets; nije u ovom fajlu) | P3 | Platform profile + NOT_APPLICABLE status | M |
| 99.51 | Ne pokriva domene iz RELEASE_GATES | narativ | **Tačno** (BankaImport samo dup :122) | P3 | Domain provideri; agregator objedinjuje | L |
| 99.52 | Dodatni nepokriveni domeni | narativ | **Tačno** (coverage gap) | P3 | Postepeno dodavati provider-e | L |
| 99.53 | Severity model nedovoljno definisan | narativ | **Tačno** (samo OK/WARN/FAIL) | P2 | Uvesti BLOCKER/FAIL/WARN/INFO/N-A/SKIPPED | M |
| 99.54 | Šta ne treba raditi | narativ | Kontekst | — | Kontekst — nije rizik | — |
| 99.55 | Preporučena ciljna arhitektura | narativ | Kontekst | — | Kontekst — nije rizik | — |
| 99.56 | Prioriteti hardeninga | narativ | Kontekst | — | Kontekst — nije rizik | — |
| 99.57 | Minimalni regression scenariji | narativ | Kontekst | — | Kontekst — nije rizik | — |
| 99.58 | Konačna procena | narativ | Kontekst | — | Kontekst — nije rizik | — |

Bilans: 46 Tačno / 2 Delimično / 10 Kontekst(-Pozitivno); hitnost: 0 P0, 0 P1, 30 P2, 18 P3. (Modul je read-only dijagnostika → nema P0/P1; sve su false-green/coverage/perf = P2 ili scope/kozmetika = P3.)

---

### FM-0089 — `modSchemaGuard.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 102.1 | Osnovni podaci (43 linije, 3 API) | Osn. podaci | Kontekst-Pozitivno | — | Kontekst — nije rizik | — |
| 102.2 | Stvarni scope = required-column+write wrapper | narativ | Kontekst | — | Kontekst — nije rizik | — |
| 102.3 | Glavni zaključak (čuva kolonu, ne identitet) | narativ | Kontekst | — | Kontekst — nije rizik | — |
| 102.4 | Šta je dobro (error kodovi 7300/7400 jedinstveni) | Pozitivno | Kontekst-Pozitivno (potvrđeno unique) | — | Kontekst — nije rizik | — |
| 102.5 | RequireColumnIndex dobra fail-fast osnova | Pozitivno | Kontekst-Pozitivno | — | Kontekst — nije rizik | — |
| 102.6 | Missing table → poruka „missing column" | narativ | **Tačno** (:10-15 GetColumnIndex=0 za obe) | P2 | Dodati RequireTable sa TABLE_MISSING kodom | S |
| 102.7 | Table case-sensitivity iz DataAccess | narativ | **Tačno** (modDataAccess.bas:79 bez vbTextCompare) | P3 | vbTextCompare u GetTable; kanon konstante mitiguju | S |
| 102.8 | RequireColumns staje na prvoj grešci | narativ | **Tačno** (:25-29 raise) | P3 | Dodati CollectMissingColumns za health/setup | S |
| 102.9 | Nemoguća/mrtva If grana | narativ | **Tačno** (:25-29 dead branch) | P3 | Zameniti sa `Call RequireColumnIndex(...)` | S |
| 102.10 | Empty ParamArray bez ugovora | narativ | **Delimično** (:25 `For 0 To -1` = tihi no-op, NIJE runtime error kako FM tvrdi) | P3 | Dokumentovati no-op ili eksplicitni guard | S |
| 102.11 | Nevalidni ParamArray elementi (CStr Null) | narativ | **Delimično** (:26 teorijski; svi calleri literali) | P3 | Nizak prioritet; validacija tipa opciono | S |
| 102.12 | RequireUpdateCell ne validira red | narativ | **Tačno** (:32-41 bez provera) | P2 | Dodati RequireValidRowIndex | S |
| 102.13 | Pozicioni rowIndex nije poslovni identitet | narativ | **Delimično** (:37 pozicioni write; AUD-003 familija, single-writer) | P2 | RequireUpdateCellByKey; = AUD-003, ne eskalirati | M |
| 102.14 | Potreban exact-row write API | narativ | Kontekst | — | Kontekst — nije rizik | — |
| 102.15 | RequireUpdateCell gubi uzrok greške | narativ | **Tačno** (:37-40 generički 7400) | P2 | Sačuvati originalni Err ili DataAccess result | M |
| 102.16 | Poruka jezički nedosledna (SR/DE) | narativ | **Tačno** (:39 „fehlgeschlagen") | P3 | Ujednačiti poruku (ASCII), dodati kontekst | S |
| 102.17 | Nema read-back verifikacije | narativ | **Tačno** (:37 samo True) | P3 | Opc. expected-value check za ID/novac/datum | M |
| 102.18 | Ne čini niz write-ova atomskim | narativ | **Tačno** (nema transakcije) | P3 | Dokumentovati oslonac na clsTransaction | S |
| 102.19 | Audit stamp best-effort iz UpdateCell | narativ | **Delimično / Nije proverivo** (StampRowAudit zavisnost) | P3 | Dokumentovati best-effort audit | S |
| 102.20 | Column cache ne invalidiran pri schema promeni | narativ | **Delimično / Nije proverivo** (mColCache zavisnost) | P3 | InvalidateSchemaCache posle Ensure*Schema | M |
| 102.21 | Ne proverava više dimenzija šeme | narativ | Kontekst (naming/scope) | — | Kontekst — nije rizik | — |
| 102.22 | Nema RequireTable | narativ | Kontekst (predlog; = 102.6) | — | Kontekst — nije rizik | — |
| 102.23 | Nema RequireAppendRow | narativ | Kontekst (predlog) | — | Kontekst — nije rizik | — |
| 102.24 | Nema structured result | narativ | Kontekst (predlog) | — | Kontekst — nije rizik | — |
| 102.25 | sourceName zavisi od discipline caller-a | narativ | **Tačno** (:20-22 bez validacije) | P3 | Standardizovati Module.Procedure format | S |
| 102.26 | Coverage širok (shared standard) | Pozitivno | Kontekst-Pozitivno | — | Kontekst — nije rizik | — |
| 102.27 | Nema dedicated regression suite | narativ | **Tačno** (fajl ne postoji) | P3 | Dodati modSchemaGuardTests | M |
| 102.28 | Šta ne treba menjati | narativ | Kontekst | — | Kontekst — nije rizik | — |
| 102.29 | Preporučeni ciljni API | narativ | Kontekst | — | Kontekst — nije rizik | — |
| 102.30 | Prioriteti | narativ | Kontekst | — | Kontekst — nije rizik | — |
| 102.31 | Regression scenariji | narativ | Kontekst | — | Kontekst — nije rizik | — |
| 102.32 | Konačna procena | narativ | Kontekst | — | Kontekst — nije rizik | — |

Bilans: 11 Tačno / 5 Delimično / 16 Kontekst(-Pozitivno); hitnost: 0 P0, 0 P1, 4 P2, 12 P3. (102.13 pozicioni write = AUD-003 familija, ne eskalira se; 102.10 delimično netačan — empty ParamArray je no-op, ne error.)

---

### FM-0090 — `modIntegritet.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 104.1 | Osnovni podaci (19 checkova, read-only) | Osn. podaci | Kontekst-Pozitivno | — | Kontekst — nije rizik | — |
| 104.2 | Stvarna odgovornost (cross-table revizija) | narativ | Kontekst | — | Kontekst — nije rizik | — |
| 104.3 | Dva režima, isti engine | Pozitivno | Kontekst-Pozitivno | — | Kontekst — nije rizik | — |
| 104.4 | Javni API | narativ | Kontekst | — | Kontekst — nije rizik | — |
| 104.5 | Check inventory A/B/C/D | narativ | Kontekst | — | Kontekst — nije rizik | — |
| 104.6 | C4 dobra remediation dijagnostika | Pozitivno | Kontekst-Pozitivno | — | Kontekst — nije rizik | — |
| 104.7 | Reuse opravdan (nasleđuje failure contract) | Pozitivno | Kontekst-Pozitivno | — | Kontekst — nije rizik | — |
| 104.8 | WriteErr ne diže totalIssues → 0 nalaza + GRESKA + info MsgBox | **Najkritičniji** | **Tačno** (:1304-1310 ne dira m_totalIssues; :59 MsgBox gleda samo total) | P1 | WriteErr uvodi ErrorCount; MsgBox INCOMPLETE kad E>0 | M |
| 104.9 | In-memory failure prikazan kao OK (Empty) | **P0** | **Tačno** (:84-85 EH→Empty; isti Empty = PASS i ERROR) | P1 | Vratiti typed IntegrityRunResult; Empty ne sme = PASS | L |
| 104.10 | Naslov overlay-a lažno zelen | narativ | **Tačno** (:90 IntegritetUkupno = m_totalIssues) | P1 | Trodelni zbir Nalazi/Greške/Status | M |
| 104.11 | Predloženi run-level rezultat | narativ | Kontekst (predlog) | — | Kontekst — nije rizik | — |
| 104.12 | Predloženi check-level rezultat | narativ | Kontekst (predlog) | — | Kontekst — nije rizik | — |
| 104.13 | „Ukupno" nije broj jedinstvenih problema | narativ | **Tačno** (:1276 sabira finding redove) | P3 | Dedup ključ CheckCode\|Entity\|Violation | M |
| 104.14 | C5 broji redove, ne duplicate grupe | narativ | **Tačno** (:551-559 svaki red) | P3 | Razdvojiti DuplicateGroups/RowsInGroups | S |
| 104.15 | Read-only tačno samo za poslovne tabele | narativ | **Tačno** (:1191-1213 piše sheet) | P3 | Precizirati opis (write diagnostic sheet/UI) | S |
| 104.16 | InitIntegritetSheet tiho ostavlja nevalidan target | narativ | **Tačno** (:1192-1200 On Error Resume Next) | P2 | Proveriti m_ws Is Nothing / Clear uspeh | S |
| 104.17 | Output write failure ne utiče na status | narativ | **Tačno** (:1217 WriteLine On Error Resume Next) | P2 | OutputSheetWritten flag; staging+atomic replace | M |
| 104.18 | ScreenUpdating se ne čuva | narativ | **Tačno** (:39,:52,:64 bezuslovno True) | P2 | Sačuvati/vratiti old vrednost (uklj. EH) | S |
| 104.19 | Aktivni sheet/selection se ne vraćaju | narativ | **Tačno** (:1211 m_ws.Activate) | P3 | activateOutput argument ili restore | S |
| 104.20 | Module-level state nereentrantan | narativ | **Delimično** (:25-30; single-thread VBA rizik nizak) | P2 | Zatvoriti u clsIntegrityRunner sa RunID | M |
| 104.21 | Nema point-in-time snapshot | narativ | **Delimično** (single-writer desktop; race nizak) | P2 | IntegritySnapshot ili cache scope | M |
| 104.22 | Ne koristi BeginTableCache | narativ | **Tačno** (:47,:93 bez scope-a) | P2 | Obmotati RunAllChecks Begin/EndTableCache | S |
| 104.23 | Soft-fail read = lažni „nema nalaza" | narativ | **Tačno** (:807 IsArray→Exit; Empty=missing i prazno) | P2 | Preflight: razlikovati praznu od nečitljive tabele | M |
| 104.24 | RequireColumnIndex tek posle uspešnog reada | narativ | **Tačno** (:807 pre :811) | P2 | Preflight schema pre svih checkova | S |
| 104.25 | Nedostaje Otkup↔Otpremnica cross-zbirna | narativ | **Tačno** (:93-113 RunAllChecks je ne zove) | P2 | Reuse canonical cross-zbirna ili dokumentovati gap | M |
| 104.26 | Header komentar zastareo („Etapa 1") | narativ | **Tačno** (:13-18) | P3 | Osvežiti header na stvarni A/B/C/D scope | S |
| 104.27 | „sve isključuju stornirane" nije doslovno | narativ | **Delimično** (kod tačan; samo doc formulacija) | P3 | Precizirati opis (context mape uključuju storno) | S |
| 104.28 | Case semantika delom dobra (vbTextCompare) | Pozitivno | Kontekst-Pozitivno | — | Kontekst — nije rizik | — |
| 104.29 | CollectBrojZbirne bez CompareMode | narativ | **Tačno** (:124 bez vbTextCompare) | P3 | Postaviti vbTextCompare kao AggByBroj | S |
| 104.30 | B6 canonical map „prvi viđeni" | narativ | **Tačno** (:888-906) | P3 | Dodati dup-normalization check za tblZbirna | S |
| 104.31 | First-seen mape kriju duplicate PK | narativ | **Tačno** (:1006,:1031,:1054 If Not Exists) | P2 | Eksplicitni duplicate-PK checkovi po tabeli | M |
| 104.32 | A1 zavisi od magic-index array | narativ | **Tačno** (:133-135 v(3)/v(0)..v(6)) | P3 | Typed rezultat ili enum indeksa iz ValidateZbirna | M |
| 104.33 | A1 tiho preskače ne-array rezultat | narativ | **Tačno** (:132-137 If IsArray bez else) | P2 | Svaki key = PASS/FINDING/ERROR, bez nestajanja | S |
| 104.34 | A2 meša integritet i business toleranciju | narativ | **Tačno / Dizajnersko** (:176-182 5%/10%) | P3 | Razdvojiti STRICT/BUSINESS_ANOMALY profil | M |
| 104.35 | Boundary strogo `>` (5%/10% granice) | narativ | **Tačno** (:176,:180 strogo >) | P3 | Dokumentovati/testirati namerni boundary | S |
| 104.36 | B7 naziv ne opisuje negativne zbirne | narativ | **Tačno** (:325 `<=0.005` uključuje negativne) | P3 | Razdvojiti EMPTY_OR_ZERO/NEGATIVE/INVALID | S |
| 104.37 | Nonnumeric → nula | narativ | **Tačno** (:841 i drugde) | P2 | Zasebna data-quality kategorija za non-num | M |
| 104.38 | A5 prijavljuje samo prvi razlog | narativ | **Tačno** (:777-782 If/ElseIf) | P3 | Dva finding-a ili lista razloga | S |
| 104.39 | C3 postojanje stavke, ne kvalitet | narativ | **Tačno** (:507 refPal.Exists) | P3 | Dokumentovati; kvalitet hvataju drugi checkovi | S |
| 104.40 | Nema severity modela | narativ | **Tačno** (svi u istom zbiru) | P2 | INFO/WARN/ERROR/CRITICAL po checku | M |
| 104.41 | Nema profila obaveznosti checkova | narativ | **Tačno** (:93-113 uvek svi) | P3 | IntegrityProfile (Core/Palletization/Processing) | M |
| 104.42 | Pokrivenost staje na preradi | narativ | Kontekst (scope; ne finance/SEF/novac) | — | Kontekst — nije rizik | — |
| 104.43 | Public bez permission guard | narativ | **Delimično** (:36 Public; read-only → nizak rizik) | P3 | Odlučiti ulogu (IZVESTAJI/Admin) | S |
| 104.44 | Environ$ Username, ne app korisnik | narativ | **Tačno** (:1204) | P3 | Čuvati AppUser (modAuth) + WindowsUser | S |
| 104.45 | Nema strukturirani run receipt | narativ | **Tačno** (nema receipt log) | P2 | Upisati INTEGRITY_RUN događaj | M |
| 104.46 | Sheet EH ne koristi centralni logging | narativ | **Tačno** (:63-66 samo MsgBox vs :84 LogErr) | P2 | Dodati LogErr u RunIntegritetProvere EH | S |
| 104.47 | Cell-by-cell output ne skalira | narativ | **Tačno** (:1293-1298 po ćeliji) | P2 | 2D niz + jedan Range.Value2 write | M |
| 104.48 | In-memory overlay bez paginga/limita | narativ | **Tačno** (:70-86, forma dodaje red-po-red) | P3 | Top-N + broj sakrivenih + filter | M |
| 104.49 | Izlaz nije deterministički sortiran | narativ | **Tačno** (dict iteracija) | P3 | Stabilan sort ključ (severity/code/entity) | S |
| 104.50 | Fiksni AutoFit A:I | narativ | **Tačno** (:1210) | P3 | Računati stvarni max broj kolona | S |
| 104.51 | Nema navigacije ka izvoru | narativ | **Tačno** (samo ID-jevi) | P3 | Hyperlink/„Otvori zapis" po finding-u | M |
| 104.52 | Nema dedicated regression suite | narativ | **Tačno** (fajl ne postoji) | P3 | Dodati modIntegritetTests | L |
| 104.53 | Minimalni P0 hardening | narativ | Kontekst (predlog) | — | Kontekst — nije rizik | — |
| 104.54 | P1 hardening | narativ | Kontekst (predlog) | — | Kontekst — nije rizik | — |
| 104.55 | P2 poboljšanja | narativ | Kontekst (predlog) | — | Kontekst — nije rizik | — |
| 104.56 | Predloženi regression scenariji | narativ | Kontekst | — | Kontekst — nije rizik | — |
| 104.57 | Šta je dobro rešeno | Pozitivno | Kontekst-Pozitivno | — | Kontekst — nije rizik | — |
| 104.58 | Najvažniji rizici (summary) | narativ | Kontekst | — | Kontekst — nije rizik | — |
| 104.59 | Konačna procena | narativ | Kontekst | — | Kontekst — nije rizik | — |

Bilans: 37 Tačno / 4 Delimično / 18 Kontekst(-Pozitivno); hitnost: 0 P0, 3 P1, 16 P2, 22 P3. (104.8/104.9/104.10 = jedan koren: false-green kad checkovi puknu — verifikovan kao najozbiljniji nalaz modula; svrstan P1 jer je dijagnostika koja laže, ne direktan gubitak podataka.)

---

Napomene za pozivaoca:
- Sva tri modula su suštinski read-only dijagnostika/write-wrapper; nijedan nalaz nije čist P0 (gubitak/korupcija). Najozbiljniji je FM-0090 104.8-104.10 (P1): `WriteErr` (`modIntegritet.bas:1304-1310`) ne inkrementira `m_totalIssues`, a `IntegritetUkupno`/MsgBox (`:59,:90`) gledaju samo taj brojač → moguće „0 neuskladjenih" uz GRESKA blokove.
- Najkorisniji verifikovani sadržajni nalaz: FM-0087 99.35 — SEF lista (`modProductionHealthCheck.bas:871`) koristi nepostojeći `SEF_CANCELLED` i propušta `SEF_REJECTED/SEF_SYNC_ERROR/SEF_TECH_FAILED` iz `modConfig.bas:659-663`.
- Ponavljajući false-green obrazac (parent OK posle child FAIL) potvrđen na dva mesta: `:951` (Google) i `:928` (soft-delete helper).
- Tri „Delimično"/korekcije FM-a: 102.10 (empty ParamArray je tihi no-op `For 0 To -1`, ne runtime error), 99.43/99.50 (zavise od modBuildGuard/modGoogleSheets, nedokazivo iz ciljnog fajla), 104.20/104.21 (reentrancy/snapshot rizik nizak na single-writer desktopu).
- Referencirani, ne-eskalirani: AUD-003 (pozicioni write — 102.12/102.13, 99.13), AUD-006 (journal — 99.5/99.11/99.13), AUD-018 (config drift — 99.5/99.12/99.13).

---

## v142 blok K2 — modRelease, modBrojevi, modParse, modHttpUtils, doccls sweep, ThisWorkbook (FM-0088/0091/0102/0104/0107/0108) [sidro origin/main v2.24.0]

Verifikacija kompletna. Svi ključni nalazi provereni protiv koda. Sastavljam audit.

---

# Audit rizik-nalaza — A2/K2 (FM-0088, 0091, 0102, 0104, 0107, 0108)

**Kod-sidro:** `origin/main` v2.24.0. **AUD mapiranja** (iz zadatka, registar je eksterni): AUD-007=DateSerial rollover, AUD-017=startup EH briše Err, AUD-034=AccessWasDenied se ne poziva / lažni STARTUP_SUCCESS posle deny-ja, AUD-037=publish guard. AUD-003 je naveden kao registrovan ali bez mapiranja u dostavljenim materijalima (`delta_intro.md:31` ga navodi bez definicije) — referenciram ga samo uz hedge gde se očigledno uklapa. **Kalibracija:** single-writer desktop; fleet-release nalazi se NE ublažavaju single-writer modelom (parcijalni release truje celu flotu), ali build je single-machine (jedan operater).

---

### FM-0088 — `modRelease.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 100.2 | Serija nezavisnih in-place upisa, ne atomski publisher | opis | **Kontekst-Pozitivno** — tačan opis (`:37-54`) | — | — | — |
| 100.3 | Nema release atomarnost | zaključak | **Tačno** — nema staging/promote | P2 | staging+`latest.json` pointer | L |
| 100.4 | Dobre osnove (Option Explicit, ext filter, binary upload) | pozitivno | **Kontekst-Pozitivno** — potvrđeno | — | — | — |
| 100.5 | **P0**: manifest se objavljuje i kad code upload padne | Kritično | **Tačno — već registrovano kao AUD-037.** `:37-47` puni `failed` ali nastavlja; `:53-54` upload `version.json` bez `If Len(failed)=0` guarda | P1 | guard pre manifest upload | S |
| 100.6 | Obrnuto: nov kod iza starog manifesta | Visok | **Tačno** — porodica AUD-037 (manifest≠fileset) | P2 | isto (atomic promote) | — |
| 100.7 | Prekid/konkurentni publish = mešavina | Visok | **Delimično** — prekid tačan; konkurentni publish je **Dizajnersko ograničenje** (single build-machine, moguće AUD-003) | P3 | remote lock (opciono) | M |
| 100.8 | Manifest nema file-listu/hash/delete | Visok | **Tačno** — `BuildManifestJson` `:69-75` ima 4 polja | P2 | dodati `files[]`+sha256 | M |
| 100.9 | Klijent koristi samo `app_version` | Visok | **Tačno** — manifest 4 polja; klijent SemVer compare | P2 | uključiti `build_sha` u odluku | S |
| 100.10 | Dva izvora identiteta (disk vs runtime konstante) | Visok | **Tačno** — `:41` disk bytes, `:69-74` runtime konst | P2 | assert source==runtime | M |
| 100.11 | `APP_VERSION` mismatch → update loop | Visok | **Tačno** — plauzibilno; nema provere | P2 | preflight source-vs-runtime | S |
| 100.12 | `BUILD_SHA` mismatch posle restore | Visok | **Tačno** — `modBuildInfo.bas:5-7` placeholder potvrđen; workflow `git checkout` | P2 | isti preflight | S |
| 100.13 | Nema placeholder/dirty zabrane | Visok | **Tačno** — `modBuildInfo` = `"0000000"`/`"0.0.0-dev"`; `PublishReleaseToDrive` ne odbija | P2 | odbij placeholder/`+dirty` | S |
| 100.14 | Nema compile gate | srednji | **Tačno** — **Dizajnersko** (ručna procedura, header `:9`) | P3 | compile receipt | M |
| 100.15 | `AssertBlankBuild` nije preduslov | srednji | **Tačno** — nema poziva | P3 | vezati receipt | S |
| 100.16 | Health/E2E nisu povezani | srednji | **Tačno** — nema poziva | P3 | gate rezultat | M |
| 100.17 | Nula fajlova = „Sve OK" | Visok | **Tačno** — `uploaded=0, failed="", okMan=True` → `:59` „Sve OK." (porodica AUD-037) | P2 | odbij `uploaded=0`+mandatory inv. | S |
| 100.18 | Broj uploadovanih ≠ kompletnost | srednji | **Tačno** — nema očekivanog inventara | P2 | canonical inventory | M |
| 100.19 | Objavljuju se svi test/dev moduli | Visok | **Tačno** — filter samo po ext `:39-40`; klijent skip samo 2 | P2 | prod exclusion/allowlist | M |
| 100.20 | Release nema profile (prod/dev) | srednji | **Tačno** | P3 | dva profila | M |
| 100.21 | `PublishReleaseToDrive` javni macro | Visok | **Tačno** — `Public Sub` bez `Option Private Module`; ali klijent nema `SRC_FOLDER` → **Dizajnersko** (single build-machine) | P3 | build-machine gate | S |
| 100.22 | Hardkodovana šifra nije zaštita | Visok | **Tačno** — `modConfig.bas:21` `RELEASE_PUBLISH_SIFRA="agrix-release"` javni const; provera samo `modAdmin.bas:284` | P3 | ukloniti/build-gate | S |
| 100.23 | Release folder = unsigned RCE trust boundary | Kritično | **Tačno** — ali granica = Drive ACL; realno P2 (single build-machine, mala flota) | P2 | potpis/hash allowlist | L |
| 100.24 | Extra remote `.bas` auto postaje update | Visok | **Tačno** — `modSelfUpdate.bas:265-276` iterira `DriveListFolder`, bez manifest allowlist | P2 | klijent čita samo manifest listu | M |
| 100.25 | Remote fajlovi se nikad ne brišu | Visok | **Tačno** — samo create-or-update | P2 | tombstone lista | M |
| 100.26 | Rename ostavlja stari+novi | Visok | **Tačno** — nema delete | P2 | signed delete | M |
| 100.27 | Lokalni self-update ne uklanja odsutne | Visok | **Tačno** — `ImportFromFolder` iterira samo skinuti folder `:301` | P2 | manifest-driven delete | M |
| 100.28 | Duplicate Drive ime nedeterministično | srednji | **Tačno** — `modDrive.bas:66` `pageSize=1`, prvi `id` `:71` | P3 | dedup provera | S |
| 100.29 | Neuspešan create+PATCH ostavlja prazan fajl | Visok | **Tačno** — `modDrive.bas:91-95` create empty, `:100-108` PATCH; PATCH fail → vrne `""` ali prazan fajl ostaje | P2 | delete-on-fail / read-after-write | M |
| 100.30 | Nema retry/backoff/content verify | Visok | **Tačno** — jedan PATCH `:100-103` | P2 | retry 429/5xx + verify | M |
| 100.31 | Hardkodovan source path (single-machine coupling) | Visok | **Tačno** — `:18` `SRC_FOLDER` const; `modVbaTools` zaseban (moguće AUD-003) | P2 | jedan build-root resolver | S |
| 100.32 | Ne proverava da je folder AgriX repo | srednji | **Tačno** — samo `FolderExists` `:30` | P3 | `.git`/marker provera | S |
| 100.33 | Enumeriše samo root fajlove | srednji | **Tačno** — `.Files` bez rekurzije `:37`; danas flat | P3 | rekurzija/expected paths | S |
| 100.34 | `.frx` se uploaduje ali ne primenjuje | Visok | **Tačno** — publisher broji `.frx` `:40`; `ImportFromFolder` ext-lista `:303` isključuje `frx`; komentar `:286` | P2 | razdvoji applied vs payload count | S |
| 100.35 | Nove forme/sheet nisu code-only | Visok | **Tačno** — `modSelfUpdate.bas:323` `st="skip"` reinstall | P2 | `requires_reinstall` | M |
| 100.36 | `modSelfUpdate` promene ne stižu floti | Visok | **Tačno** — `SKIP_MODULES` `:38` | P2 | `minimum_updater_version` | M |
| 100.37 | `modVbaTools` nije runtime update | srednji | **Tačno** — `SKIP_MODULES` `:38`; **Prihvaćeno** (dev tool) | P3 | payload klasifikacija | S |
| 100.38 | Nema updater protocol verzije | Visok | **Tačno** | P2 | `manifest_schema_version` | S |
| 100.39 | Nema `requires_reinstall` | Visok | **Tačno** | P2 | detekcija+polje | M |
| 100.40 | Nema compile provere posle klijent merge-a | srednji | **Tačno** | P2 | post-merge compile check | M |
| 100.41 | Jedan globalni kanal, nema canary | srednji | **Tačno** — `REL_FOLDER_ID` jedan | P3 | stable/canary | L |
| 100.42 | Nema immutable prethodni release/rollback | Visok | **Tačno** | P2 | versioned folderi | L |
| 100.43 | Pre-update backup ne rešava fleet integritet | kontekst | **Kontekst-Pozitivno** — tačna nijansa | P3 | — | — |
| 100.44 | Rezultat nije strukturiran/trajan | srednji | **Tačno** — `Sub`+MsgBox `:56-61` | P3 | `PublishReleaseResult`+receipt | M |
| 100.45 | Ne loguju se file ID-jevi | srednji | **Tačno** — `Len(...)>0` bul. signal `:41` | P3 | mapa ime→id | S |
| 100.46 | Temp manifest se ne briše; `WriteReleaseTextFile` nema EH close | nizak | **Tačno** — nema `Kill`; `:77-82` Open/Print/Close bez EH | P3 | `Kill`+EH close | S |
| 100.47 | Manifest nema JSON/schema validaciju | nizak | **Tačno** — ručni concat `:69-75` | P3 | parse-back+required | S |
| 100.48 | Release datum ≠ publish datum | nizak | **Tačno** — `BUILD_DATE` iz commita | P3 | `published_at` | S |
| 100.49 | Nema tenant/workbook compat metapodataka | srednji | **Tačno** | P3 | schema-level polja | M |
| 100.50 | Nema dependency/mandatory-module provere | Visok | **Tačno** | P2 | mandatory inventory gate | M |
| 100.51–100.57 | Predloženi protokol/folderi/manifest/preflight/result/prioriteti/regresija | predlog | **Predlog (ne nalaz)** — razuman ciljni ugovor; van minimal-delta scope-a | — | usvojiti inkrementalno | L |
| 100.58 | Konačna procena | sažetak | **Kontekst** — saglasan | — | — | — |

**Bilans FM-0088:** Sve činjenične stavke **Tačno** (nula opovrgnutih); 2 kalibrisane na **Dizajnersko ograničenje** (100.7 konkurentni publish, 100.14 compile gate) i 1 na **Delimično** (100.7). Jedini pravi **P1** je publish-guard 100.5/100.6/100.17 = **AUD-037** (minimal fix S: `If Len(failed)=0 And uploaded>0` pre `version.json` upload-a + odbij `uploaded=0`). Ostalo P2/P3 — realan arhitektonski dug (immutable versioned package + hash manifest), ali single build-machine/mala flota spuštaju hitnost ispod fleet-P0 tona FM-a. `.frx` broj (100.34) i `SKIP_MODULES` posledice (100.36) su tačni „applied≠uploaded" nalazi.

---

### FM-0091 — `modBrojevi.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 106.2–106.5 | Kontekst, format, API semantika | opis | **Kontekst-Pozitivno** — tačno; „Generate*" ime sugeriše jaču garanciju (106.5 tačno) | P3 | rename → `Suggest*`/typed | S |
| 106.6 | Dobro rešeno (max iz brojeva, storno rezervisan) | pozitivno | **Kontekst-Pozitivno** — potvrđeno | — | — | — |
| 106.7 | Generator predloga, ne atomski allocator | zaključak | **Tačno** — read-then-return `:93-99` | P2 | save-time guard | M |
| 106.8 | **P0**: nema atomsku rezervaciju | Kritično | **Tačno kao arhitektonska granica** — **Dizajnersko/Prihvaćeno** za single-writer (verovatno pokriveno AUD-003); **P1 tek u SWMR** | P2 | backend seq+CAS (SWMR) | L |
| 106.9 | Otvorena forma → predlog zastari | Visok | **Tačno** — single-writer smanjuje verovatnoću | P2 | re-check pre save | S |
| 106.10 | **P0**: `MaxSeqFromTable` na grešku → 0 | Kritično | **Tačno** — `:389-391` EH→0; `RequireColumnIndex` (`modSchemaGuard.bas:13`) baca, ali EH guta → fail-open | P2 | typed `ALLOCATION_UNAVAILABLE` | M |
| 106.11 | **P0**: `GenerateBrojPrijemnice` EH vraća `1/ddmmyy` | Kritično | **Tačno** — `:203` EH vraća validan-looking duplikat | **P1** | EH → `""` (hard fail), ne fallback broj | S |
| 106.12 | Caller prijemnice nastavlja bez statusa | Visok | **Tačno** — `modAutoHladnjaca.bas:150` → `SavePrijemnica_TX :161`, bez provere/unique | P2 | proveri `Len=0` pre write | S |
| 106.13 | **P0**: `BrojZbirneExists` na grešku → False | Kritično | **Tačno** — `:419-421` EH→default False (UNKNOWN→FREE) | P2 | tri-state / retry | S |
| 106.14 | **P0**: mirror detekcija fail-open | Kritično | **Tačno** — `IsStanicaMirrorVozac :287` `On Error Resume Next`→False | P2 | tri-state UNKNOWN | S |
| 106.15 | Mirror uslov proverava samo postojanje stanice | Visok | **Delimično/Dizajnersko** — `:289-290` samo `LookupValue tblStanice`; ali komentar `:285` to i definiše kao contract | P3 | dokumentuj+strict | S |
| 106.16 | **P0**: remote failure = max 0 | Kritično | **Tačno** — `MaxSeqFromGoogleSheet` vraća 0 na SVIM putanjama (`:429-481`); ali samo suggestion put (`checkRemote`) | P2 | strict mode / typed remote status | M |
| 106.17 | Potrebna dva režima (best-effort/strict) | predlog | **Predlog (ne nalaz)** | — | — | — |
| 106.18 | **P0**: paralelni generator `modMasterSync` | Kritično | **Tačno** — `modMasterSync.bas:2887-2928` row-count (`seq=seq+1`), ne max-seq | **P1** | delegirati `modBrojevi.MaxSeqFromTable` | M |
| 106.19 | Direktan duplicate scenario (rupe) | Kritično | **Tačno** — potvrđeno: `1/190726`+`1/190726-3` → 2 reda → `seq=1+2=3` → `1/190726-3` duplikat | **P1** | isto (ukloniti row-count) | M |
| 106.20 | `modMasterSync` nema globalni ZBR bump | Visok | **Tačno** — `:2927` samo `ApplyMirrorPrefix`, bez `BrojZbirneExists` petlje (koju `SuggestNextBroj :106-109` ima) | P1 | delegirati canonical | M |
| 106.21 | Dokumentacija precenjuje centralizaciju | Visok | **Tačno** — dva generatora + derivacije | P3 | canonical matrica | S |
| 106.22 | Toggle gasi samo UI suggestion | srednji | **Tačno** — `IsAutoBrojDokumenta` (`modConfig.bas:962`) proveravan samo u `SuggestNextBroj :46` | P3 | razdvoji policy imena | S |
| 106.23 | `SuggestNextBroj` validira samo neprazan ID | srednji | **Tačno** — `:51` samo `Len(Trim)=0` | P3 | validiraj numeric>0 | S |
| 106.24 | `FormatBroj` prihvata nulti namespace | srednji | **Tačno** — `0/ddmmyy` prolazi regex `^\d+/...` | P3 | odbij entity=0 | S |
| 106.25 | `ExtractNumericFromEntityID` konkatenira sve cifre | srednji | **Tačno** — `:214-217` petlja svih cifara | P3 | strict suffix parser | M |
| 106.26 | Nema overflow guard | Visok | **Tačno** — `:222` `CLng(digits)`, public, bez lokalnog EH → overflow ruši caller-a | P2 | `TryExtract` sa bounds | S |
| 106.27 | `ExtractSeqFromBroj` previše tolerantan | srednji | **Tačno** — `abc/xyz→1`, `x/y-999→999` (`:227-253`) | P3 | strict varijanta za allocator | S |
| 106.28 | Regex proverava sintaksu, ne semantiku | srednji | **Tačno** — `:260` bez datum/suffix validacije | P3 | kalendarska provera | S |
| 106.29 | Generički validator ne podržava `S` | Visok | **Tačno** — `IsValidBrojFormat :260` `^\d+...` bez S; `modMasterSync.bas:2797-2804` zaseban `IsValidBrojZbirneFormat` `^S?\d+...` = format drift | P2 | `ValidateBroj(kind,val)` | M |
| 106.30 | `ApplyMirrorPrefix` slab contract | srednji | **Tačno** — `:296-301` samo `Left$="S"`, bez trim/format (`s1..→Ss1..`) | P3 | parse-normalize | S |
| 106.31 | `MaxSeqFromTable` case-sensitive entity | Visok | **Tačno** — `:371` `CStr(...)=entityID` exact → trailing space/case preskočen | P2 | canonical normalizacija | S |
| 106.32 | Nevalidan datum reda tiho ignorisan | Visok | **Tačno** — `:374-376` `CDate` pod `On Error Resume Next` → red ne ulazi u max | P2 | strict blok na corrupt | S |
| 106.33 | Datum kolona i datum-u-broju mogu protivrečiti | srednji | **Tačno** — nema invariant provere | P3 | invariant check | M |
| 106.34 | Lokalni OTK fallback ne proverava remote | Visok | **Tačno** — `GenerateBrojDokumenta :137` skenira samo `tblOtkup` | P2 | dok. import-order contract | S |
| 106.35 | OTP VBA-only ≠ single-instance | srednji | **Tačno** — **Prihvaćeno** (single-writer) | P3 | SWMR backend | — |
| 106.36 | PRJ prefiks `1` hardkodovan | srednji | **Tačno/Dizajnersko** — `:198` `FormatBroj("1",...)`; komentar `:180-185` definiše pravilo | P3 | formalizuj policy | S |
| 106.37 | Remote header lookup strog fail-soft | srednji | **Tačno** — `FindHeaderIndexInData :518` exact compare→0 | P3 | `REMOTE_SCHEMA_INVALID` | S |
| 106.38 | Hardkodovan `Sheet1` | srednji | **Tačno** — `:442` | P3 | schema contract | S |
| 106.39 | Cache čuva prazan spreadsheet ID | Visok | **Tačno** — `:503` upisuje i `""` → session-long negative cache | P2 | ne keširati prazno | S |
| 106.40 | Cache nema TTL/generation | srednji | **Tačno** — key samo `sheetName :503` | P3 | key=folder+tenant+gen | S |
| 106.41 | `ClearSpreadsheetIDCache` nema caller-e | srednji | **Tačno** — grep: samo deklaracija `:305`, nula poziva | P3 | zvati na config-change | S |
| 106.42 | Remote suggestion čita ceo sheet | srednji | **Tačno** — `ReadSheetData(...,"Sheet1") :442` | P3 | server-side seq | L |
| 106.43 | Revers ceo `tblAmbalaza`; missing kol→0 | srednji | **Tačno** — `MaxSeqReversAmbalaza :324` `GetColumnIndex`; `iBroj=0→Exit` (default 0); pozitivno: uključuje storno | P3 | typed unavailable | S |
| 106.44 | Dvocifrena godina = business format | kontekst | **Kontekst-Pozitivno** — tačno (interni ID postoji) | P3 | — | — |
| 106.45 | Business date contract mora biti eksplicitan | srednji | **Tačno** — funkcije primaju `Date`, bez policy | P3 | dokumentuj | S |
| 106.46–106.48 | Nema `AllocationResult`; ciljni API; SWMR model | predlog | **Predlog (ne nalaz)** — razuman pravac | P3 | typed rezultat | L |
| 106.49 | Manual unos takođe zahteva unique guard | Visok | **Tačno** — nema save-time unique provere | P2 | finalni exact-unique guard | M |
| 106.50–106.52 | Jedna canonical impl.; logging≠receipt; nema test suite | predlog/nalaz | **Tačno** (106.51 logging, 106.52 nema `modBrojeviTests` — potvrđeno) | P2 | dedic. regresija | M |
| 106.53–106.56 | Šta ne menjati / prioriteti / regresija / procena | meta | **Kontekst** — saglasan | — | — | — |

**Bilans FM-0091:** Sve činjenične stavke **Tačno** (nula opovrgnutih). Ključna kalibracija: FM markira 8 nalaza kao „P0", ali single-writer desktop spušta većinu fail-open readova (106.10/13/14/16) na **P2** i atomsku rezervaciju (106.8) na **Dizajnersko/AUD-003** (P1 tek u SWMR). Dva prava **P1** danas: **106.11** (`GenerateBrojPrijemnice` EH → validan-looking `1/ddmmyy` duplikat — fix S) i **106.18–106.20** (`modMasterSync.GenerateBrojZbirne` row-count generator pravi duplikat na rupama — realan bug i single-writer, fix M delegiranjem). Format drift 106.29 i overflow 106.26 potvrđeni. Pozitivno: max-iz-brojeva algoritam i storno-rezervacija su solidni.

---

### FM-0102 — `modParse.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 121.1 | Uloga (4 javne funkcije, cross-cutting) | opis | **Kontekst-Pozitivno** — tačno | — | — | — |
| 121.2 | Pozitivna arhitektura (TryParse, NBSP/RSD/kg) | pozitivno | **Kontekst-Pozitivno** — potvrđeno | — | — | — |
| 121.3 | **P0/P1**: `DateSerial` prihvata nepostojeće datume | Kritično | **Tačno — već registrovano kao AUD-007.** `:71-72` bez re-check-a; `31.02` normalizuje | (AUD-007) | `Day/Month/Year=input` provera | S |
| 121.4 | Locale-zavisnost / `IsDate` dvosmislenost | Visok | **Tačno** — `:50-52` `IsDate/CDate` pre fallback-a; `03/04/2026` per-mašina | P2 | canonical `dd.mm.yyyy` | M |
| 121.5 | Dvocifrena godina asimetrična | srednji | **Tačno** — `:69` `Y<100→2000+Y` bez dok. contract-a | P3 | dokumentovan pivot / odbij | S |
| 121.6 | Jedan separator semantički dvosmislen | Visok | **Tačno** — `:114-119` uvek decimalni; `1.234→1,234` na SR (factor-1000) | P2 | contract: zabrani grouping / strict | M |
| 121.7 | Preširoko uklanjanje jedinica | Visok | **Tačno** — `:90-91` `Replace RSD/kg` svuda; `1kg2→12`, `RSD100RSD→100` | P2 | samo leading/trailing token | S |
| 121.8 | `ByRef result` ostaje star na neuspehu | srednji | **Tačno** — `:11,29` Exit bez reset-a | P2 | `result=0` na početku | S |
| 121.9 | `TryParseLong` granice/naziv | srednji | **Tačno** — `:29-33` Double-first, neg odbijen, tol `0.000001` | P3 | rename `TryParseNonNegativeLong` | S |
| 121.10 | Fail-closed bez observability-ja | srednji | **Tačno** — sve greške → False bez razloga | P3 | typed parse reason | M |
| 121.11 | Caller površina (široka) | kontekst | **Tačno** — potvrđeno 11 caller-fajlova (frmOtkup, frmStammdaten, frmDokumenta, frmBankaExportPregled, modOtkupBlok, modSEFMapper/Validator, modDokumenta, modBankaImportParserPdfToText…) | — | regresija pre promene contract-a | — |
| 121.12 | Test pokrivenost nedostaje | srednji | **Tačno** — nema dedic. matrice | P2 | offline test matrica | M |
| 121.13–121.14 | Prioriteti / ocena | meta | **Kontekst** — saglasan | — | — | — |

**Bilans FM-0102:** Sve **Tačno**. 121.3 = **AUD-007** (referencirano, ne re-analizirano). Ostalo: 121.6 (single-separator factor-1000) i 121.7 (unit-strip u sredini) su najkonkretniji finansijski rizici — obe **P2** sa jeftinim fix-om (S). 121.4 locale nedeterminizam P2. Nijedan nalaz nije opovrgnut; FM ovde precizan i pošten (i sam kod komentariše neka ograničenja).

---

### FM-0104 — `modHttpUtils.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 124.1 | Uloga (UrlEncode/JsonEscape, dedup) | opis | **Kontekst-Pozitivno** — tačno (`:15-25` uklanja 4 duplikata, UTF-8 bug fix) | — | — | — |
| 124.2 | `UrlEncode` tehnički dobro izveden | pozitivno | **Kontekst-Pozitivno** — potvrđeno (`:52` UTF-8 bytes, BOM strip, RFC3986 unreserved) | — | — | — |
| 124.3 | Windows/ADODB dep; `Utf8Bytes` bez cleanup | srednji | **Tačno** — `:103-142` nema `On Error`/`stream.Close` na grešci | P2 | EH+close+re-raise | S |
| 124.4 | `JsonEscape` nije kompletan JSON encoder | Visok | **Tačno** — `:91-96` samo `\ " CRLF/CR/LF`; nema `\t \b \f \u00XX`; kod to i priznaje `:84-87` | P2 | escape svih U+0000–001F | S |
| 124.5 | Novi redovi se normalizuju na `\n` | kontekst | **Kontekst-Pozitivno** — tačno i poželjno | — | — | — |
| 124.6 | Caller odgovoran za navodnike | kontekst | **Kontekst-Pozitivno** — `:82` komentar tačan | — | dok. primer | S |
| 124.7 | Nema builder-a (O(n²)) | nizak | **Tačno** — `:69-73` `result=result&`; trenutni calleri kratki | P3 | ne za bulk | — |
| 124.8 | Test pokrivenost | srednji | **Tačno** — grep: nema `Test_UrlEncode`/`Test_JsonEscape` | P2 | offline matrica | M |
| 124.9–124.10 | Prioriteti / ocena | meta | **Kontekst** — saglasan | — | — | — |

**Bilans FM-0104:** Sve **Tačno**. FM markira 124.4 kao P1; kalibrišem na **P2** — kanonski helper treba da bude ispravan, ali trenutni SEF/Google payload-i nemaju sirove control-znakove (kod svesno dokumentuje granicu), pa nije produkcioni blocker. Jedini realni budući rizik: control-znak iz copy/paste/banke → nevalidan JSON. Fix jeftin (S). `UrlEncode` je stvarno poboljšanje (UTF-8 bug fix za srpska slova).

---

### FM-0107 — Empty worksheet-class sweep (`*.doccls`)

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| — | Svih 12 `.doccls` = samo standardni metadata blok, bez event/procedure logike | nalaz | **Tačno** — potvrđeno: svaki fajl tačno 9 linija (`VERSION/BEGIN/MultiUse/END/Attribute×5`), nula procedura; nema `Worksheet_*` | P2 | — | — |
| — | Arhitektonsko značenje (nema skrivenog sheet-event sloja) | zaključak | **Kontekst-Pozitivno** — tačno | — | — | — |
| — | Rizik: repo noise + nejasna `CodeName↔sheet` mapa (`Sheet14/15/16/20/23`) | P2 | **Tačno** | P3 | manifest CodeName↔ime↔tabela | S |
| — | Preporuke (ne brisati, CI import-gate, tanki eventi ako dođu) | predlog | **Predlog (ne nalaz)** — razumno | — | — | — |

**Bilans FM-0107:** Potvrđeno trivijalno tačno — metadata-only workbook-object exporti, nema P0/P1. Jedini realni nalaz je P3 higijena (CodeName mapa). Klasifikacija „KEEP, DOCUMENT, DO NOT TREAT AS FUNCTIONAL" je ispravna.

---

### FM-0108 — `ThisWorkbook.doccls`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 108.1 | Tanak lifecycle adapter (2 handlera, delegira) | opis | **Kontekst-Pozitivno** — tačno (`:10-56` Open, `:58-84` BeforeClose) | — | — | — |
| 108.2 | **P1**: lažni `VBA_STARTUP_SUCCESS` | Visok | **Tačno** — `:24-34` emit bezuslovan; `StartApp` normalno izlazi na 4 mesta: `modMain.bas:38` (AccessGate=deny, **= AUD-034**), `:44` (UpdateGate), `:50-53` (self-update scheduled), `:59-64` (login failed) — sva 4 `Exit Sub` | P2 | `StartAppResult`; success samo za `Started` | M |
| 108.3 | **P1**: `InitApp` može pasti, `StartApp` nastavlja | Visok | **Tačno** — `InitApp` `:220-222` `Resume CleanUp` (ne re-raise); `ValidateAllTables` je `Private Sub` (`:287-309`) — samo MsgBox, ne vraća False/ne baca; `:211-212` `m_Initialized=True` bezuslovno; `StartApp :32` ne proverava rezultat | P2 | `InitApp` fail-closed; ne `m_Initialized=True` uz missing tabele | M |
| 108.4 | Monitoring identitet hardkodiran (`"Operator"`) | srednji | **Tačno** — `:12,29` `userId:="Operator"`, `correlationId:="VBA-STARTUP"` konstanta | P3 | device/session pre login; RunID | S |
| 108.5 | Cleanup orphan lockova nevidljiv caller-u | srednji | **Tačno** — `:20-21` `CleanupOrphanedLocks` pod `On Error Resume Next`, bez `Err` provere (vezano za **AUD-017** — startup EH briše Err) | P3 | `STARTUP_DEGRADED_LOCK_CLEANUP` warn | S |
| 108.6 | `BeforeClose` nema strukturiran rezultat release-a | srednji | **Tačno** — `:65-76` best-effort, nema monitoringa nereleased lock-a | P3 | warn StationID+razlog | S |
| 108.7 | Shutdown state ne garantovano restauriran (`mIsShuttingDown`) | srednji | **Tačno** — `modMain.bas:257-260` EH ne vraća `mIsShuttingDown=False` | P3 | reset u EH; dok. best-effort | S |
| 108.8 | Pozitivni nalazi (kratak, hvata fatal, oslobađa lock) | pozitivno | **Kontekst-Pozitivno** — potvrđeno | — | — | — |
| 108.9–108.10 | Prioriteti / ocena (typed startup rezultat) | meta | **Kontekst** — saglasan | — | — | — |

**Bilans FM-0108:** Sve **Tačno**. Registrovano (referencirano, ne re-analizirano): **AUD-017** (startup EH briše Err — dodiruje 108.5), **AUD-034** (`AccessWasDenied` se ne poziva / lažni STARTUP_SUCCESS posle deny-ja — poklapa se sa slučajem (1) u 108.2). **Novo/šire od AUD-034**: 108.2 pokriva i slučajeve (2) UpdateGate, (3) self-update scheduled, (4) login failed — telemetrijski „success" iako UI nije pokrenut → **P2** (kvari operacionu telemetriju, ne funkciju). 108.3 je stvaran fail-open: `ValidateAllTables` kao `Private Sub` + bezuslovni `m_Initialized=True` znači da app može biti „inicijalizovan" uz nedostajuće tabele → **P2**. Ostalo P3 (observability). Najbolji minimal-delta: typed `StartAppResult` između `ThisWorkbook` i `modMain` (rešava 108.2+108.3 zajedno).

---

## Zbirni bilans (A2/K2)

- **Ukupno verifikovano:** 6 FM unosa, ~150 stavki/podsekcija. **Nula činjenično opovrgnutih** — FM je i u ovom bloku faktografski precizan; sve „Netačno" = 0.
- **Već registrovano (referencirano):** AUD-007 (FM-0102 121.3), AUD-017 + AUD-034 (FM-0108), AUD-037 (FM-0088 100.5/100.6/100.17 publish-guard). AUD-003 referenciran uz hedge na koncept-atomarnost (106.8 / 100.7/100.31) — tačan opseg nije u dostavljenim materijalima.
- **Pravi P1 (fix S–M), nezavisni od SWMR:**
  1. **FM-0091 106.11** — `GenerateBrojPrijemnice` EH (`modBrojevi.bas:203`) vraća validan-looking `1/ddmmyy` duplikat → EH treba `""`. (S)
  2. **FM-0091 106.18–106.20** — `modMasterSync.GenerateBrojZbirne` (`:2887-2928`) row-count generator pravi duplikat na rupama; delegirati canonical `MaxSeqFromTable`. (M)
  3. **FM-0088 100.5** = **AUD-037** — guard pre `version.json` upload-a (`modRelease.bas:53`). (S)
- **Kalibracija naniže:** većina FM „P0/Kritično" tona spušta se na **P2/P3** zbog (a) single-writer desktop-a (modBrojevi fail-open readovi 106.10/13/14/16, atomarnost 106.8), (b) single build-machine / male flote (modRelease security/canary 100.21–100.23, 100.41), (c) dokumentovanih dizajn-granica (106.15, 106.36, 124.4 kontrolisani payload). Fleet-integritet nalazi (100.24–100.35) ostaju legitimni P2 arhitektonski dug jer se NE ublažavaju single-writer modelom.
- **Format napomena:** FM-0102/0104/0107 su čisti (nula opovrgnutih, uglavnom P2/P3 hardening); FM-0088/0091 nose najveći realan dug ali sa jasnim minimal-delta ulaznim tačkama.

Relevantni fajlovi: `modRelease.bas`, `modBrojevi.bas`, `modParse.bas`, `modHttpUtils.bas`, `ThisWorkbook.doccls` (svi u `src-vba/`, verifikovano protiv `origin/main` v2.24.0); cross-ref dokazi: `modMasterSync.bas:2887-2928`, `modSchemaGuard.bas:5-18`, `modMain.bas:16-309`, `modDrive.bas:54-113`, `modSelfUpdate.bas:38,262-336`, `modConfig.bas:21`, `modAdmin.bas:284`, `modBuildInfo.bas:5-7`, `modAutoHladnjaca.bas:150-161`.

---

## v142 blok K3 — modMasterSync, modGoogleSyncSmokeTests (FM-0092/0093) [sidro origin/main v2.24.0]

Verifikacija kompletna. Sve ključne tvrdnje proverene protiv koda. Sledi audit.

---

## AUDIT — FM-0092 & FM-0093 vs modMasterSync.bas / modGoogleSyncSmokeTests.bas (origin/main v2.24.0)

Napomena o kontekstu: single-writer desktop; jedina multi-writer površina je Google Sheet (GAS/PWA). Zato su „race" tvrdnje realnije na GS strani (writeback), a slabije na lokalnom append-u (tblOtkup/tblZbirna). Već registrovano (referencira se, ne eskalira): AUD-002 = 108.7, AUD-018 = 108.40, AUD-001 = zaseban parser u modGoogleSheets (≠ 108.46).

### FM-0092 — `modMasterSync.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 108.1–108.6 | Osnovni podaci / uloga / tokovi / „šta je dobro" | narativ/poz. | **Kontekst-Pozitivno** | Prihvaćeno | Bez izmene; opis tačan (`Option Explicit`, strict headeri, `ClientRecordID`, row-TX potvrđeni) | — |
| 108.7 | OTK batch TX može ostaviti `Synced>Master` bez master reda | P0 | **Tačno** — `287-312`: outer snapshot→`Core(False)`→`RollbackTx`; `Core` vraća False na `totalErrors>0` (`213-235`), writeback već upisan u `ImportOneOTKSheet` (`1356-1362`) | **= AUD-002 (već registrovano)** | — | — |
| 108.8 | OTK dva consistency modela po entry-pointu | P0 (drift) | **Tačno** — `Core` direktno (`163-182`) ostavlja uspešne redove; `_TX` (`298-301`) rollback-uje sve. Ista jezgra, druga semantika | P1 | Ukloniti outer-TX (isto kao AUD-002 fix); `_TX` = tanak wrapper nad `Core` bez table-snapshot rollbacka | M |
| 108.9 | VOZ put pravilnije modelovan (no-outer-TX) | pozitivno | **Kontekst-Pozitivno** — `2167-2179` eksplicitni komentar + `ImportZbirneFromPWA_TX` (`2160`) ne pravi snapshot | Prihvaćeno | Zadržati; primeniti isti model na OTK | — |
| 108.10 | Treba inbox/outbox umesto pseudo-distributed TX | P0 (arh.) | **Dizajnersko ograničenje** — validan pravac, nije defekt po sebi | P2 | `tblSyncInbox` receipt (Domain+CRID unique) tek posle stabilizacije 108.7/108.8 | L |
| 108.11 | Writeback vezan za fizički `rowNum` (F/B/T) | P0 | **Tačno** — `1729`,`1737` (OTK), `2839`,`2845`,`2853` (VOZ); `rowNum` iz read-a, bez CRID precondition/read-back | P1 (GS je multi-writer) | Pre upisa read-back ćelije A i assert `A==očekivani CRID`; bolje GAS `ack` po CRID-u | M–L |
| 108.12 | PWA lock ne rešava identity-safe writeback | P0 (analiza) | **Tačno** — potpora 108.11; lock ne pokriva ručni GS edit/sort ni javne entry-pointe | P1 | Vidi 108.11 | — |
| 108.13 | dedupe check i append nisu atomski | P0 | **Delimično** — `IsDuplicateInMaster`(`1451`)→`AppendRow`(`1623`) jeste neatomski, ali lokalni append je single-writer desktop; trigger je 2 Excel instance | P2 (P0 tek za SWMR) | CRID unique-guard pri append-u; za sada dokumentovati single-instance pretpostavku | M |
| 108.14 | Duplicate check ne prijavljuje postojeći `count>1` | visok | **Tačno** — `1470-1473` / `2560-2563` vraćaju True na prvom match-u | P2 | Kada `count>1` → hard integrity error (kao `RequireSingleMasterSyncRow`), ne „Duplicate" | S |
| 108.15 | `TryUpdateVozacID` zaobilazi hardening; True i na neuspeh | visok | **Tačno** (realan bug) — `1773-1798`: `GetColumnIndex`/`UpdateCell` (ne Require*), rezultat `UpdateCell` se ne proverava, `=True`+log „Updated" bezuslovno (`1790-1793`) → GS dobija `Synced>Master`, `VozacID` prazan | P1 | Proveriti Boolean `UpdateCell`; `=True` samo ako write uspeo; koristiti `RequireUpdateCell` | S |
| 108.16 | Duplicate status ne vraća postojeći master ID | visok | **Tačno** — OTK `1319/1322` i VOZ `2360` upisuju `Duplicate` bez `OtkupID/ZbirnaID/BrojZbirne` | P2 | Na duplicate vratiti postojeći ID u ServerRecordID (idempotent replay) | M |
| 108.17 | Invalidan datum → današnji datum | P0 | **Tačno** (OBA puta) — OTK `1547-1550`, VOZ `2598-2600`: `On Error Resume Next / CDate / If Err Then Date` | P1 | Strict parse; neuspeh → row `SyncError:Invalid Datum`, ne `Date` | S |
| 108.18 | OTK validacija suviše uska | visok | **Delimično / Dizajnersko** — `1381-1449` proverava koop/vrsta/kol/cena/amb; nema ownership/klasa/datum/parcela. Defense-in-depth iza GAS auth-a | P2 | Dodati proveru `Klasa` enum, stanica-aktivna, koop∈stanica (min. skup) | M |
| 108.19 | Sheet ownership (`OTK-ST-x` ⇒ `OtkupacID=ST-x`) se ne proverava | visok | **Tačno** — `ImportOneOTKSheet` prima `sheetName` ali row-import ga ne poredi (`1253+`); isto VOZ | P2 | Assert `OtkupacID`/`VozacID` == entitet iz imena sheet-a | S |
| 108.20 | Poslovni broj samo regex | visok | **Tačno** — `IsValidBrojFormat`(OTK `1597`) / `IsValidBrojZbirneFormat`(`2639`) — sintaksa, ne semantika (prefiks/datum/zauzetost) | P2 | Semantička provera prefiksa vs entitet + `ddmmyy` vs datum | M |
| 108.21 | PWA `BrojZbirne` bez uniqueness pre append-a | P0 | **Tačno** — `2626-2644`: regex pa direktan `AppendRow`; ne zove `BrojZbirneExists`/`SuggestNextBroj`. Dovoljan 1 batch sa 2 ista broja | P1 | Pre append-a `BrojZbirneExists(broj)` → konflikt = `SyncError` | M |
| 108.22 | Fallback ZBR generator paralelni i pogrešan | P0 | **Tačno** — `GenerateBrojZbirne`(`2887-2924`) BROJI redove (`seq=count+1`) → za `1/ddmmyy`+`1/ddmmyy-3` daje postojeći `-3`; kanonski `SuggestNextBroj(KIND_ZBR)` radi max-suffix + `BrojZbirneExists` bump-loop (modBrojevi `93-110`). Latentno (samo prazan-broj fallback) | P1 | Zameniti telo sa `SuggestNextBroj(KIND_ZBR, vozac, datum)` | M |
| 108.23 | Duplirani formatter/validator već driftovali | visok | **Tačno** — `IsValidBrojZbirneFormat` regex `^S?\d+/...` vs modBrojevi `IsValidBrojFormat` `^\d+/...` (razlika: `S` prefiks, `257-263`); `ExtractNumericVozacBroj`(`3007`) duplira `ExtractNumericFromEntityID` | P2 | Prošriti kanonske u modBrojevi (opc. `S`), obrisati lokalne kopije | M |
| 108.24 | `KulturaID` fallback pravi orphan | visok | **Tačno** — `1583-1584`: `KulturaID = VrstaVoca & "-" & SortaVoca` ako lookup padne; vrednost ne mora postojati u `tblKulture` | P2 | Nepoznata kultura → `SyncError:Unknown culture` umesto sintetičkog ID-a | S |
| 108.25 | Positional `AppendRow` = schema-order rizik | visok | **Tačno / Dizajnersko** — `1615-1623` (OTK), `2674-2677` (ZBR) veliki positional nizovi; poznat rizik iz CLAUDE.md | P3 | Prelazak na upis po imenu (`UpdateCell`/`GetColumnIndex`) za osetljiva polja | L |
| 108.26 | Ambalažni dvojni ledger zavisi od slabo validirane stanice | srednji | **Tačno** — `1629-1630` Izlaz/Ulaz; `stanicaID` iz koop-mappinga ili fallback `otkupacID` (`1575-1580`) | P2 | Vezati za 108.19 ownership-guard | S |
| 108.27 | Auto-otpremnica grupiše po nedovoljnom ključu | P0 | **Tačno** — `668-672` ključ `Stanica\|Datum\|Vozac\|Klasa`; metadata sa prvog reda (`739-744`), kol/amb sabrani (`717-718`) → mešane vrste/cene/ambalaže u 1 otpremnicu | P1 | U ključ dodati `VrstaVoca\|SortaVoca\|Cena\|TipAmbalaze`; različita metadata = odvojene otpremnice | M |
| 108.28 | Auto-otpremnica nema granicu ture | visok | **Tačno / Dizajnersko** — model nema `TuraID`; isti ključ spaja dve vožnje | P2 | Zahteva `TuraID`/shipment grain u domenu | L |
| 108.29 | `otpAll` preload je mrtav kod | srednji | **Tačno** — `692-697` učitava `otpAll`+`colOtpSt/colOtpDat`, nigde se ne koriste (broj ide preko `GenerateBrojOtpremnice` `724`) | P3 | Obrisati `otpAll`/`colOtpSt`/`colOtpDat` | S |
| 108.30 | Malina stamping menja SVE aktivne prazne `VozacID` | visok | **Delimično** — `816-844` prolazi ceo aktivni `tblOtkup`; nije skopiran na run/PWA/sezonu. U malina modu je `VozacID:=StanicaID` namerni semantik | P2 | Ograničiti na tekući import (npr. `SyncSource=PWA` + prazan Otpremnica), ne ceo history | M |
| 108.31 | Malina stamping ne garantuje mirror vozača | srednji | **Delimično** — helper (`816`) ne zove `EnsureVozacMirrorForStanica`; mirror obezbeđuje `modAutoHladnjaca`/`modMalina` na drugom mestu (mitigacija zavisi od orchestrator redosleda) | P2 | Ili pozvati `EnsureVozacMirrorForStanica` u stamping-u, ili dokumentovati garanciju u orchestratoru | S |
| 108.32 | Auto-zbirna iz otpremnice nema conflict guard | visok | **Tačno** — `895-986`: `brZbirne=ApplyMirrorPrefix(...)`→`SaveZbirna_TX` bez provere globalne zauzetosti/konflikta | P2 | `BrojZbirneExists` pre `SaveZbirna_TX` | M |
| 108.33 | Backfill ćutke zadržava konfliktni `BrojZbirne` | visok | **Tačno** — `BackfillOtkupBrojZbirneByOtpremnica` `1006`: update samo ako `cur=""`; različit postojeći ostaje bez prijave | P2 | `cur<>""` i `cur<>novi` → log/`SyncError` konflikt | S |
| 108.34 | VOZ payload može linkovati proizvoljne OTK CRID-eve | P0 | **Tačno** — `LinkZbirnaToOtkupAndOtpremnica` `2701-2771`: po CRID-u samo `RequireSingleMasterSyncRow`; nema provere vozač/datum/stanica/već-u-drugoj-zbirnoj | P1 | Validirati membership (isti vozač/datum, nije u drugoj zbirnoj) pre linkovanja | M |
| 108.35 | Postojeće veze se prepisuju bez conflict politike | P0 | **Tačno** — `RequireUpdateCell ... COL_OTK_BROJ_ZBIRNE`(`2759`) i `LinkOtpremnicaToBrojZbirneStrict`(`1894`) upisuju bez „prazno ILI isto" guarda | P1 | Guard: dozvoli upis samo ako prazno ili identično; inače konflikt | M |
| 108.36 | Prazan `OtkupRecordIDs` ipak commit-uje zbirnu | visok | **Tačno** — `2712-2715` log+Exit, pa `tx.CommitTx`(`2463`) upisuje standalone zbirnu | P2 (zavisi od ugovora) | Ako je zbirna = skup otkupa → prazna lista = `SyncError` | S |
| 108.37 | Nema quantity reconciliation zbirne vs otkupi | srednji | **Tačno** — nigde `Σ Otkup.kol vs KolI+KolII` pre commit-a | P2 | Reconcile pre commit-a ili eksplicitno prepustiti `modIntegritet` | M |
| 108.38 | PWA zbirna grain (combined I/II) ≠ lokalni per-class | srednji | **Delimično** — kod potvrđen (`2664-2677` jedan red, `Klasa="I/II"`); „downstream mora podržati oba" je inferencija | P2 | Dokumentovati oba grain-a; verifikovati downstream po-klasi izveštaje | L |
| 108.39 | VOZ row failure gubi precizan razlog | srednji | **Tačno** — `2387` piše generički `SyncError:Import/link failed`; tačan `Err` samo `Debug.Print`(`2487-2491`) | P2 | Typed reason code iz `ImportVOZRow_RowTX` u status | S |
| 108.40 | VOZ Drive listing nema paginaciju | P1 | **Tačno** — `FindVOZSheets` `2243-2246` samo `pageSize=100`, bez `nextPageToken` | **= AUD-018 (već registrovano)** | — | — |
| 108.41 | OTK/VOZ implementacije driftovale | strukturno | **Tačno** — potvrđeni parovi (listing paginacija Da/Ne, 2 writeback helpera, dupli JSON parser) | P2 | Zajednički sync engine tek posle contract hardeninga | L |
| 108.42 | Provisioning proverava postojanje, ne ispravnost | srednji | **Tačno** — `469-474`: `existingID` po imenu → skip; bez header/tab/kolone/schema-version provere | P2 | Za existing sheet validirati header schema | M |
| 108.43 | Neuspešan header write ostavlja poison spreadsheet | P0/P1 | **Tačno** — `476-494`: `CreateSpreadsheet` pa `WriteSheetData`; na fail header-a `newID` se NE trash-uje → sledeći run ga vidi kao existing | P1 | Trash na fail header-a, ili temp-ime pa rename posle uspešnog write-a | M |
| 108.44 | Postojeći 22-kol sheet nema migracioni put | visok | **Tačno** — `BuildOTKOperationalHeaders_` 23 kol (`544-568`), provisioning skip existing, validator traži kol 23 | P2 | Repair/dopuna postojećih sheetova na 23-kol schemu | M |
| 108.45 | Header count guard pogrešan | srednji | **Tačno** — `1177` `< 22` (poruka „Expected=22"), a `GS_BROJ_DOKUMENTA=23` traži kol 23 (`1207`); 22-kol prolazi count, pada na pristupu kol 23 (fail-safe ali poruka pogrešna) | P3 | Promeniti guard na `< 23` i poruku „Expected=23" | S |
| 108.46 | Ručni JSON parser krhak (Drive lista) | srednji | **Tačno** — `ExtractJsonValueAt` `1133-1147` InStr/prvi navodnici; bez escaped/unicode. Za `id,name` rizik mali. Različit parser od AUD-001 | P3 | Ne širiti; dugoročno pravi JSON parser | M |
| 108.47 | Nema retry/backoff | srednji | **Tačno** — svi `WinHttp` pozivi direktni bez retry (429/5xx/401-refresh/timeout) | P2 | Centralni retry wrapper (429/5xx/timeout, 1× refresh na 401) | M–L |
| 108.48 | Writeback bez chunking-a | srednji | **Tačno** — `1704-1743` / `2827-2859` jedan JSON body za sve update-e | P3 | Deterministički chunk-ovi + per-chunk receipt | M |
| 108.49 | Full-sheet scan degradira | srednji | **Tačno** — `ImportOneOTKSheet` čita ceo sheet; `IsDuplicateInMaster` re-čita celu tabelu po redu | P3 (scale) | CRID set učitati jednom po batch-u; server-side status filter | M |
| 108.50 | `skipped` metrika neupotrebljiva | nizak | **Tačno** — `1349`/`2398` broji svaki non-`Synced` red (stari `Synced>Master`, `Duplicate`...) | P3 | Odvojeni brojači po statusu | S |
| 108.51 | Monitoring meša VOZ kao `PWA-OTKUP` | srednji | **Tačno** — `Monitor_MasterSync*` hardkoduje `entityID="PWA-OTKUP"`, `correlationId="MASTERDATA-SYNC-PWA"` (`3047-3048`,`3064-3065`,`3081-3082`), zove ga i VOZ core | P3 | Parametrizovati domain (OTK/VOZ) u monitor helperima | S |
| 108.52 | Nema `SyncRunID` | srednji | **Tačno** — statičan correlationId, bez per-run identiteta | P3 | Uvesti `SyncRunID` (GUID) kroz ceo ciklus | M |
| 108.53 | `mLastPWAFatalSyncError` shared mutable | nizak | **Tačno / Prihvaćeno** — modul-level Boolean, resetuje se na startu core-a; OK za single-thread | P3 | Strukturiran result dugoročno | M |
| 108.54 | Remote status state machine neformalizovan | srednji | **Tačno** — stringovi bez dozvoljenih tranzicija | P3 | Definisati enum + dozvoljene tranzicije | L |
| 108.55 | `ServerRecordID` preopterećen semantički | srednji | **Tačno** — writeback upisuje `OtkupID/ZbirnaID` u kol B (`ServerRecordID`) | P2 | Odvojiti `GasRecordID` / `MasterRecordID` | M |
| 108.56 | Geo pull dobra identity/rollback zaštita | pozitivno | **Kontekst-Pozitivno** — `3245-3275` dup source/local = hard fail; 1 TX; identity-based update; prazno ne briše | Prihvaćeno | Zadržati | — |
| 108.57 | Geo source nije domenski validiran | srednji | **Tačno** — `GeoUpdateFieldIfNeeded` `3408-3438` samo non-empty+different; bez lat/lng range/polygon/enum | P2 | Type-aware validacija po koloni pre upisa | M |
| 108.58 | Geo merge ne podržava namerno brisanje | srednji | **Tačno / Dizajnersko** — `3425` prazno ne briše (štiti od gubitka) | P3 | Eksplicitni delete-marker kad zatreba | M |
| 108.59 | Geo string-compare pravi write churn | nizak | **Tačno** — `3429` `StrComp` binarno nad `CStr` (44.123 vs 44,123 vs 44.1230) | P3 | Type-aware normalizacija po koloni | M |
| 108.60 | Missing Google parcele van Boolean rezultata | srednji | **Tačno** — `3260-3268` broji/loguje, funkcija i dalje `True`(`3311`) | P3 | Vratiti structured counts (updated/skipped/missing) | S |
| 108.61 | Cohesion hotspot 3.446 linija | strukturno | **Tačno** (činjenično) | P2 (advisory) | Split po failure granicama tek posle contract hardeninga | L |
| 108.62 | Dokumentacija (runbook) delimično zastarela | nizak | **Nije proverivo statički** (kod ≠ docs) — kod jeste „prihvati broj iz reda, generiši lokalno samo ako prazan" (`2626-2636`) | P3 | Uskladiti runbook (ko alocira/validira/authority) | S |
| 108.63 | Smoke fixture 22-kol schema drift | P0 | **Tačno** — ukršteno sa FM-0093/110.6 | P1 | Vidi FM-0093 | S |
| 108.64 | Smoke fixture header `Količina` ≠ validator | — | **Tačno** — ukršteno sa 110.6 | P1 | Vidi FM-0093 | S |
| 108.65 | Test coverage uži od rizika modula | visok | **Tačno** — nema VOZ/writeback-failure/reorder/concurrency scenarija | P2 | Dodati VOZ + failure/recovery suite (vidi FM-0093) | L |
| 108.66 | Procena stila (novi strict + stari soft) | narativ | **Kontekst-Pozitivno** — tačan rezime | Prihvaćeno | — | — |
| 108.67–108.70 | Predlozi: structured result / receipt / writeback zaštita / otpremnica politika | predlog | **Dizajnersko ograničenje** — validni pravci | P1–P2 | Prioritet: writeback CRID-precondition (108.11) + otpremnica group-key (108.27) | L |
| 108.71–108.74 | P0/P1/P2 lista + regression scenariji | rezime | **Tačno** (mapira potvrđene nalaze) | — | Koristiti kao backlog; ispravne stavke 1,3,4,5,7,8,9,10 | — |
| 108.75 | Konačna procena | zaključak | **Kontekst-Pozitivno** — verno | Prihvaćeno | — | — |

Bilans FM-0092: Od ~55 sadržajnih nalaza — **Tačno** ~40, **Delimično** 5 (108.13/108.18/108.30/108.31/108.38), **Kontekst-Pozitivno** 5 (108.5/108.9/108.56/108.66/108.75), **Dizajnersko** 4 (108.10/108.25/108.28/108.58 + predloзi), **Nije proverivo** 1 (108.62). Nijedan **Netačan**. Već registrovano: 108.7=AUD-002, 108.40=AUD-018. Realni implementacioni bug-ovi (ne samo arh.): **108.15** (`TryUpdateVozacID` True na neuspeh), **108.17** (datum→danas, oba puta), **108.22** (row-count ZBR generator pravi duplikat), **108.29** (mrtav `otpAll`), **108.43** (poison sheet), **108.45** (guard `<22`/kol 23). Najveća poslovna korupcija van AUD-002: **108.27** (mešane vrste/cene u 1 otpremnicu) i **108.35** (prepis postojećih veza). P1 fokus: 108.11/108.15/108.17/108.21/108.22/108.27/108.34/108.35/108.43.

---

### FM-0093 — `modGoogleSyncSmokeTests.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 110.1–110.4 | Osnovni podaci / podela suite / inventory | narativ | **Kontekst-Pozitivno** — opis tačan (2 javna entry-pointa, 2 suite) | Prihvaćeno | — | — |
| 110.5 | Šta je dobro (real roundtrip, token se ne štampa) | pozitivno | **Kontekst-Pozitivno** — `Option Explicit`(`3`), komentar „Never print token"(`101`), realni create/write/read/find | Prihvaćeno | Zadržati real-Google roundtrip | — |
| 110.6 | OTK fixture ne odgovara 23-kol schema ugovoru | P0 | **Tačno** — `BuildOTKFixtureData` `541` `1 To 22` (nema `BrojDokumenta`); `571` `"Koli"&ChrW(269)&"ina"`=`Količina`, validator traži `Kolicina` (`1199`, binarni `StrComp`) → happy-path pada pre importa | P1 | Fixture na 23 kol + header `Kolicina`; graditi iz kanonskog header providera | S |
| 110.7 | Root uzrok: header definisan na više mesta | P0 (arh.) | **Tačno** — header u `BuildOTKOperationalHeaders_`(modMasterSync `543`), `ValidateOTKSheetHeader`(`1159`), GAS, i `BuildOTKFixtureData`(`539`) | P1 | `modPWASyncSchema` provider koji koriste provisioning+validator+fixture | M |
| 110.8 | Suite failures se ne propagiraju caller-u | P0 | **Tačno** — `RunGoogleSyncSmokeSuite` je `Sub` bez povratne vrednosti; svaki test EH→`LogGoogleSmokeFail`→End Sub; fatal EH (`47-52`) log+cleanup+završi | P0 | `..._Core() As Boolean = (m_Failed=0)`; javni Sub `Err.Raise` ako False | M |
| 110.9 | Konkretan E2E false-green | P0 | **Tačno** — `modE2EReleaseGate.E2E_RunVbaSuite` `71-82`: `Application.Run procName`(`75`)→`E2E_Pass`(`77`) osim ako digne grešku; suite hvata sve → PASS i kad interno FAIL (poruka `77` doslovno: „Verify suite summary ... for PASS counters") | **P0** | Zavisi od 110.8: gate mora čitati Boolean/structured result, ne completion | M |
| 110.10 | Potreban izvršni test contract | P0 (predlog) | **Dizajnersko ograničenje** — validan predlog (`_Core As Boolean` / `TestSuiteResult`) | P0 (uz 110.8/110.9) | Min. varijanta: Boolean + `Err.Raise`; gate proverava `Succeeded` | M |
| 110.11 | Rollback celih produkcionih tabela (SWMR) | P0 (SWMR) | **Delimično** — `RunMasterSyncSmokeSuite` `371-380` snapshot+`RollbackTx` cele `tblOtkup`/`tblAmbalaza`; na single-writer desktopu bezbedno, opasno pri paralelnom writeru/SWMR | P2 (P0 za SWMR) | Cleanup samo po test-ownanim ID-jevima pod lock-om; izolovan test workbook | M–L |
| 110.12 | Bezbedniji cleanup model (owned IDs) | P0 (predlog) | **Dizajnersko ograničenje** — validan pravac | P2 | `TestRunID` + brisati samo svoje `OtkupID/AmbID` | M |
| 110.13 | Rollback ne vraća sve side-effekte | visok | **Tačno** — `AppendRow` (modDataAccess `209-215`) radi `StampRowAudit`+`WriteJournalRow`+`InvalidateTableCache`+`gKpiDirty=True`; `RollbackTx` vraća samo ListObject snapshot → orphan journal red ostaje | P2 (recovery noise, ne korupcija) | Izolovati destruktivni test od produkcionih tabela; ili purge journal po test-CRID-u | M |
| 110.14 | Nested transaction semantika | visok | **Tačno** (strukturno) — outer test TX + inner `ImportRowToTblOtkup_RowTX` commit (`1489-1501`) + writeback + outer rollback = split-brain obrazac | P2 | Vezano za 110.11/110.12 izolaciju | M |
| 110.15 | Fixture nije determinističan + pogrešan config ključ | visok | **Tačno** — „prvi red" iz koop/kultura/stanica (`548-554`, `640`); `GetConfigValue("DefaultSorta")`(`551`) — kanonski je `CFG_DEFAULT_SORTA="DEFAULT_SORTA_VOCA"` (modConfig `575`, editor modPodesavanja `75`) → uvek prazno→`"Default"` | P1 | Fiksan deterministički tuple; ključ `CFG_DEFAULT_SORTA` | M |
| 110.16 | Fixture može sakriti ownership grešku | visok | **Tačno** — koop (`548`) i stanica/`OtkupacID` (`554`) biraju se nezavisno; import izvodi StanicaID iz koopa pa `OtkupacID` fallback (modMasterSync `1575-1580`) → mismatch prolazi | P2 | Izabrati koop koji pripada izabranoj stanici | M |
| 110.17 | Assertions preslabe za business import | visok | **Tačno** — `413-455`: proverava `imported=1`, `errors=0`, `Synced>Master`, ServerRecordID≠"" ; ne proverava `ServerRecordID==OtkupID`, tačan `count=1`, polja, 2 amb reda, netaknutost drugih | P2 | Exact-receipt assertions (ID jednakost, polja, 2 amb reda) | M |
| 110.18 | Duplicate test ne dokazuje idempotency | visok | **Tačno** — `462-494`: `imported=0`, `skipped>=1`, `Duplicate`; bez count pre/posle, bez existing-ID, bez amb-side-effect | P2 | Dodati count-invariant + existing-ID + „nema novog amb reda" | S |
| 110.19 | Missing-CRID test ne potvrđuje „no side effects" | srednji | **Tačno** — `501-526`: `imported=0`, `errors>=1`, `SyncError*`; bez row-count invariant / prazan ServerRecordID / tačan razlog | P2 | Assert `tblOtkup`/`tblAmbalaza` count nepromenjen + tačan reason | S |
| 110.20 | VOZ/Zbirna potpuno nepokrivena | visok | **Tačno** — `TestHook_ImportOneVOZSheet` i `BuildVOZFixtureData` ne postoje nigde (grep prazan); modul ima samo `BuildOTKFixtureData` | P1 | Dodati VOZ hook (modMasterSync) + `BuildVOZFixtureData` + happy/dup/link testove | L |
| 110.21 | Glavni failure prozor (writeback fail posle commit-a) netestiran | visok | **Tačno** — nema fault injection (401/403/429/timeout/partial/disconnect); sve happy real-API | P1 | Fault-injection sloj oko writeback-a (mock HTTP status) | L |
| 110.22 | Write-back identity (row moved) netestiran | visok | **Tačno** — nijedan test ne pomera red pre writeback-a niti proverava CRID na target redu | P2 | Test: pomeri red/promeni CRID pre writeback-a → mora biti odbijen (vezano uz 108.11 fix) | M |
| 110.23 | Google smoke = environment test, ne unit | srednji | **Tačno** — zavisi od realnog tokena/foldera/mreže/kvote | P3 | Preimenovati/razdvojiti Offline/Online/Destructive suite | S |
| 110.24 | Cleanup nije potvrđen read-back-om | srednji | **Tačno** — `TrashGoogleDriveFile` `218-253` True na bilo koji 2xx (`248`); jedan `m_TestSpreadsheetID` | P2 | Read-back `trashed=true` + exact-name search prazan; resource registry | S |
| 110.25 | `RunID` nije concurrency-safe | srednji | **Tačno** — `263`/`399` `Format$(Now,"yyyymmddhhnnss")` (1s rezolucija) | P3 | GUID / machine+session+nonce | S |
| 110.26 | Nema reentrancy guard | srednji | **Tačno** — modul-level `m_Total/m_Passed/m_Failed/m_RunID/m_TestSpreadsheet*` (`21-27`) pretpostavljaju 1 aktivnu suite | P3 | `m_RunInProgress` guard; dugoročno `clsTestRunContext` | S |
| 110.27 | Counter model nije broj testova | srednji | **Tačno** — `AssertTrue/Equals`→`LogGoogleSmokePass`(`307/317`) + test-end pass (`66/83/102/157/189`) + cleanup(`206`) svi u `m_Total` | P3 | Odvojiti TestCases vs Assertions vs CleanupFailures metrike | M |
| 110.28 | Pogrešan naziv završnog summary-ja | srednji | **Tačno** — `RunMasterSyncSmokeSuite`→`EndGoogleSmokeRun`(`383`) prikazuje „Google Sync Smoke Suite PASS/FAIL"(`277/281`) | P3 | Per-suite `SuiteName` u summary-ju | S |
| 110.29 | EH može sakriti neuspešan rollback | srednji | **Tačno** — `386-392` `On Error Resume Next` pa `RollbackTx`/cleanup/`EndGoogleSmokeRun`; neuspeh rollbacka nije zaseban rezultat | P2 | Razdvojiti `LocalCleanupSucceeded`/`RemoteCleanupSucceeded`; FAIL ako bilo koji cleanup nepotvrđen | S |
| 110.30 | Test komentari zastareli | srednji | **Tačno** — `14-15` „Does NOT write to tblOtkup", a `RunMasterSyncSmokeSuite` piše u `tblOtkup`/`tblAmbalaza` (preko `TestHook`→`AppendRow`) | P2 | Zameniti header komentarom „DESTRUCTIVE INTEGRATION TEST — writes to production tables" | S |
| 110.31 | Predlog nove organizacije modula | predlog | **Dizajnersko ograničenje** — validno (pure/online/destructive split) | P2 | Posle P0 popravki | L |
| 110.32 | Predlog canonical OTK fixture-a | predlog | **Dizajnersko ograničenje** — dobar (named setter iz header array-a) | P1 (uz 110.6/110.7) | Fixture iz `GetOTKOperationalHeaders()` | M |
| 110.33–110.35 | P0 test matrice (OTK/VOZ/Google negative) | predlog | **Tačno** (mapira realne rupe) | P1–P2 | Backlog za novu suite; prioritet: OTK 23-kol happy, VOZ happy, writeback-fail | L |
| 110.36 | Prioriteti hardening-a (P0/P1/P2) | rezime | **Tačno** | — | P0: 110.6/110.8/110.9/110.11 | — |
| 110.37 | Šta NE raditi (ne brisati real roundtrip) | savet | **Kontekst-Pozitivno** — ispravno | Prihvaćeno | Zadržati real-Google, samo izolovati/označiti | — |
| 110.38 | Ocena trenutnog stanja (tabela) | rezime | **Tačno** — „P0 false green", „VOZ ne postoji", „SWMR neprihvatljiva" potvrđeni | — | — | — |
| 110.39 | Konačni zaključak (3 razloga) | zaključak | **Tačno** — fixture drift + FAIL ne propagira + whole-table rollback svi potvrđeni | Prihvaćeno | — | — |

Bilans FM-0093: Od ~30 sadržajnih nalaza — **Tačno** ~24, **Delimično** 1 (110.11 — SWMR-uslovljeno), **Dizajnersko** 4 (110.10/110.12/110.31/110.32), **Kontekst-Pozitivno** 4 (110.1–110.5/110.37). Nijedan **Netačan**. Najkritičnije i najčvršće potvrđeno: **110.9 (P0 E2E false-green)** — `E2E_RunVbaSuite` marker PASS na `Application.Run` bez izuzetka, a suite guta sve greške; sama poruka priznaje da ne čita brojače. To je jače od fixture drifta (110.6): čak i kad fixture eksplicitno padne, gate prijavljuje PASS. Realni bug-ovi: **110.6** (22-kol + `Količina`), **110.8** (Sub bez rezultata), **110.15** (ključ `DefaultSorta` umesto `DEFAULT_SORTA_VOCA`), **110.20** (VOZ hook ne postoji), **110.30** (lažni „ne piše u tblOtkup" komentar). P0 lanac za popravku: 110.8 (izvršni rezultat) → 110.9 (gate ga čita) → 110.6/110.7 (canonical fixture) → 110.11 (izolacija pre SWMR).

Napomena o preseku: 108.63/108.64 (FM-0092) = 110.6/110.7 (FM-0093) — isti schema-drift, jedan fix. Ni jedan nalaz nije zahtevao izmenu koda (READ-ONLY audit). Sve `fajl:linija` citate možeš direktno otvoriti u `src-vba/` (verifikovano protiv `origin/main` v2.24.0).

---

## v142 blok K4 — modMalina, modSledljivost, frmSledljivost, frmAgrohemija, modAgrohemija (FM-0096/0105/0111/0112/0113) [sidro origin/main v2.24.0]

I now have enough verification across all five files plus the cross-module dependencies. Let me compile the audit.

Verification complete. Below is the audit of every FM finding against the actual v2.24.0 code.

---

### FM-0096 — `modMalina.bas`

Kontekst: single-writer desktop; mirror se kreira i kroz backfill (`modMalina`), i kroz `frmStammdaten` hook, i kroz `modAutoHladnjaca`. Verifikovao sam sve cross-module tvrdnje (`modBrojevi` L286-290, `modAutoHladnjaca` L112-119, `modMasterSync` L816-843, `frmStammdaten` L2354-2355, `modBusinessFlowProTests` L1532-1544).

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 116.5/116.9 | „Ensure" je create-only shadow bootstrap, ne pravi mirror (postojeći red se nikad ne usklađuje; L39-41 Exit odmah) | Zaključak | **Tačno** | P2 | Preimenovati u `Reconcile*` i dodati update grane; kratkoročno bar dokumentovati semantiku | M |
| 116.6 | Dve komponente različito definišu mirror identitet: `modMalina` ne proverava stanicu, `modBrojevi.IsStanicaMirrorVozac` (L289-290) proverava samo `tblStanice` → orphan-shadow (A) i missing-shadow (B) stanja | **P0** | **Tačno** | **P1** | Jedan canonical helper `IsManagedStationMirror(id)` koji traži par `tblStanice`+`tblVozaci`; koristiti ga u oba modula | M |
| 116.7 | `EnsureVozacMirrorForStanica` ne čita `tblStanice` (ne potvrđuje da stanica postoji/aktivna/jedinstvena) | (P0 grupa) | **Tačno** | P1 | Reconcile varijanta sama čita canonical red stanice po `stanicaID` | S |
| 116.8 | Idempotency check nije exact-one guard (L39 `LookupValue` vraća prvi; ne broji redove) | — | **Tačno** | P2 | `FindRows(...).Count`: 0→create, 1→ok, >1→integrity error | S |
| 116.10 | Nova stanica: mirror dobija `telefon=""` (`frmStammdaten` L2354-2355 šalje `""`; kasniji Ensure preskoči) | Konkretan bug | **Tačno** | P2 | Proslediti `Trim$(txtField3.value)` (Kontakt) umesto `""`; reconcile bi ionako popravio | S |
| 116.11 | Izmena stanice (`btnIzmeni_Click`) ne poziva Ensure/reconcile → trajni drift | — | **Tačno** | P2 | Dodati reconcile hook posle commit-a izmene stanice | S |
| 116.12 | Nema deaktivacije/reaktivacije mirror-a; status stanice se ne propagira | — | **Tačno** | P2 | Reconcile sinhronizuje `Aktivan` iz stanice | M |
| 116.13 | `BackfillVozacMirrorsForMalina` ne filtrira neaktivne stanice (L93-101, nema `Aktivan` provere) | — | **Tačno** | P2 | Politika aktivna→aktivan / neaktivna→neaktivan, ista u create/backfill | S |
| 116.14 | Boolean `False` meša 5 stanja (mode off/prazan id/postoji/AppendRow=0/greška) | — | **Tačno** | P2 | Enum/typed rezultat ili bar razdvojiti „error" od „no-op" | M |
| 116.15 | `AppendRow=0` je silent failure (L61 samo `>0`→True, inače tih False) | — | **Tačno** | P1 | `If AppendRow<=0 Then Err.Raise` sa kontekstom | S |
| 116.16 | Ensure `EH` guta grešku (L68 samo `LogErr`, bez re-raise); caller sa `On Error` ne dobija signal | — | **Tačno** | P1 | Re-raise u EH (kao što backfill radi na L111) | S |
| 116.17 | Backfill može prijaviti `created=0` kad su svi redovi pali (Ensure guta per-row grešku pa ne stiže do backfill EH) | — | **Tačno** | P1 | Structured backfill rezultat (Scanned/Created/Failed…) | M |
| 116.18 | Backfill parcijalno commituje bez RunID/failure liste/residue audita | — | **Tačno** (Dizajnersko za idemp. model) | P3 | Vratiti brojila i listu failures | M |
| 116.19 | Check-then-append nije SWMR bezbedan (dve instance mogu dodati isti `VozacID`) | — | **Dizajnersko ograničenje** (kontekst = single-writer; FM sam kvalifikuje „u SWMR arhitekturi") | Prihvaćeno/P3 | Ako multi-writer postane realnost: post-append exact-one provera; inače prihvatiti | S→L |
| 116.20 | `modAutoHladnjaca` ignoriše rezultat Ensure-a i bezuslovno `vozacID=stanicaID` (L114-118) → dokument može dobiti FK bez `tblVozaci` reda | — | **Tačno** | P1 | Proveriti rezultat / reconcile pre `vozacID=stanicaID`; blokirati lanac ako mirror nije obezbeđen | M |
| 116.21 | `modMasterSync.StampVozacFromStanicaForMalina` (L836-838) upisuje `StanicaID` u prazan `VozacID` bez Ensure/exact-one provere | — | **Tačno** | P1 | Pozvati reconcile pre stamping-a ili posle-batch health provera | M |
| 116.22 | Malina toggle nema atomsku migraciju (ključ se menja, backfill ručno) | — | **Tačno** (Dizajnersko) | P2 | Vezati aktivaciju za preflight+backfill | M |
| 116.23 | Nema markera da je vozač shadow (`SourceStanicaID`/`MirrorManaged`); tip se zaključuje iz ID-a | — | **Tačno** | P2 | Dodati marker kolone pri reconcile | M |
| 116.24 | `Naziv→Ime`, `Mesto→Prezime` semantička zloupotreba (backfill L88-101; Ensure L52-53) | P2 | **Tačno** | P2 | `DisplayName` polje ili read-model union; ne dupliranje u person kolone | L |
| 116.25 | `Aktivan` opciono (L58-59): ako kolona fali, red se svejedno kreira bez statusa | — | **Tačno** | P3 | Za canonical šemu tretirati `Aktivan` kao obavezan (warning ako fali) | S |
| 116.26 | Dupli komentarisani `Attribute VB_Name` (L1-2) | Nefunkcionalno | **Tačno** | P3 | Ukloniti komentarisanu liniju (export normalizacija) | S |
| 116.27 | Test proverava samo postojanje reda + drugi poziv False (L1538-1544) | — | **Tačno** | P2 | Proširiti na exact-one/polja/mode-off/missing-col | M |
| 116.28 | Test dozvoljava orphan mirror — `MIR_ST` (L1532) se ne kreira u `tblStanice` pre Ensure → validira vozača bez stanice | — | **Tačno** | P2 | Seed-ovati `tblStanice` red pre Ensure; proveriti `IsStanicaMirrorVozac` | S |
| 116.29 | Fiksni `ST-MIRTEST-90001` ostaje trajno u `tblVozaci` (nema cleanup) | — | **Tačno** | P2 | Disposable fixture ili cleanup na kraju testa | S |
| 116.30-116.34 | Regression matrica, `ReconcileVozacMirrorForStanica` API, `TransportActor` model, health invariant | P1/P2 predlozi | **Dizajnersko ograničenje** (predlozi, ne bug-ovi) | P2/P3 | Prihvatljiv ciljni hardening; nije P0 preduslov | L |
| 116.2/116.4/116.36 | Pozitivno: `Option Explicit`, self-gated `IsMalinaMode` (L33), prazan `sid` odbačen (L36), `ReDim` na `ListColumns.Count`+`RequireColumnIndex` (L50-53), reuse `AppendRow` | Pozitivno | **Kontekst-Pozitivno** | — | Zadržati | — |

Bilans FM-0096: Sve konkretne stavke su činjenično tačne protiv koda (verifikovano i 5 cross-module tvrdnji). Ključna zamerka na FM-ovu kalibraciju: **P0 je precenjeno za single-writer desktop**. 116.19 (SWMR race) je van modela (jedan pisac); orphan-iz-izmišljenog-ID-a (116.7/116.28) nema production caller-a koji fabrikuje ID (svi prosleđuju realan `StanicaID` iz `tblStanice`/`frmStammdaten`). Realno štetni lanac je **116.6 + 116.20 + 116.21** (missing-shadow → `S`-broj + orphan FK na dokumentu) → moja preporuka P1. Najjeftiniji visoko-vredni potez: jedan canonical `IsManagedStationMirror` + re-raise u Ensure EH (116.6/116.16, napor S-M).

---

### FM-0105 — `modSledljivost.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 122.3 | `TraceByZbirna` filtrira po `OtpremnicaID` (L540-544), ne po canonical `tblOtkup.BrojZbirne` → otkup sa `BrojZbirne` ali praznim `OtpremnicaID` se izostavlja | **P1** | **Tačno** | **P1** | Dodati direktan `BrojZbirne` prolaz kroz `tblOtkup`, obeležiti `OtpremnicaID` status | M |
| 122.4 | Trace nema `isComplete`/broj očekivanih vs pronađenih/warning | P1 | **Tačno** | P1 | Typed rezultat (`Succeeded/IsComplete/Rows/Warnings`) | M |
| 122.5 | `Empty` maskira greške — `GetUnlinkedOtkupi`/`TraceByZbirna` na EH vraćaju `Empty` (L403, L603), isto kao „nema podataka" | P1 | **Tačno** | P2 | Razdvojiti `NoData` od `Failed` (typed rezultat / out-param) | M |
| 122.6 | Auto-link `0` ne kaže razlog (nema kandidata/ambiguous/već povezano/nevalidan datum/greška) | — | **Tačno** | P2 | Vratiti brojila umesto golog `Long` | M |
| 122.7 | Nevalidan datum → prazan ključ (`BuildAutoLinkKey` L272) → red tiho ispada, batch prijavi `Linked` bez `InvalidKeyCount` | — | **Tačno** | P2 | `InvalidKeyCount` brojač + evidencija | S |
| 122.8 | Prazni `StanicaID/VozacID/Klasa` su validni segmenti ključa (nema non-empty zahteva) → link po nepotpunom identitetu | — | **Tačno** | P2 | Odbiti auto-link kad strict segment fali (osim legacy politike) | S |
| 122.9 | `BrojZbirne` poređenje nenormalizovano u trace-u: auto-link `UCase$+Trim$` (L282) vs trace `CStr(...)=brojZbirne` sirovo (L464) | — | **Tačno** | P1 | `UCase$(Trim$())` + `vbTextCompare` u trace-u, isto kao auto-link | S |
| 122.10 | Lookup referencijalne greške → prazna polja bez orphan indikatora (L566-585) | — | **Tačno** | P2 | Orphan marker kad `LookupValue` prazan a ID postoji | S |
| 122.11 | N+1 lookup (per-otkup `LookupValue` × 7 na `tblKooperanti`/`tblParcele`) | P2 | **Tačno** | P2 | `BuildLookupDict` jednom pre petlje | M |
| 122.12 | Krhak positional 14-kolonski `Variant` bez named contract-a (L419-426: 8=BPG,10=GGAP,12=ParcelaID,14=Površina) | — | **Tačno** | P3 | Typed UDT/klasa za trace red | M |
| 122.13 | Zbirna/otpremnica nisu zasebni trace čvorovi (samo otkup redovi) → naziv širi od contract-a | — | **Tačno** (Dizajnersko) | P3 | Preimenovati ili proširiti u pravi graf-trace | L |
| 122.14 | `AutoLinkOtkupOtpremnica_TX` vraća `0` i na failure (L82) = isto kao uspešan no-op | — | **Tačno** | P2 | Typed rezultat sa `Succeeded`+brojila | M |
| 122.15 | Monitoring `userId:="Operator"` hardkodovan (L34, L72) | P2 | **Tačno** | P3 | Stvarni auth/session ID | S |
| 122.16 | Postojeći link se nikad ne revalidira (L190 preskače otkup sa `OtpremnicaID`) | — | **Tačno** | P2 | Odvojen `ValidateExistingTraceLinks` health-check | M |
| 122.17 | Nema zaštite od reused `BrojZbirne` kroz generacije (grupiše sav tekst) | — | **Tačno** (Dizajnersko) | P2 | Dugoročno `ZbirnaID` kao canonical ključ | L |
| 122.18 | `GetUnlinkedOtkupi` ne vraća razlog (7 polja, bez `BrojZbirne`/klase/candidate count) | — | **Tačno** | P2 | Proširiti izlaz razlogom + brojem kandidata | S |
| 122.2 | Pozitivno: `ExcludeStornirano`, `RequireColumnIndex`, `GetUniqueAutoLinkTarget` exact-one (L308), OtkupID exact-one pre upisa (L243-247), TX wrapper+rollback | Pozitivno | **Kontekst-Pozitivno** | — | Zadržati | — |

Bilans FM-0105: Svi nalazi tačni. Dva su suštinska (sledljivost/compliance): **122.3** (trace zavisi od pomoćnog linka, ne od canonical `BrojZbirne` — realno moguć nepotpun trag) i **122.9** (nekonzistentna normalizacija broja u istom modulu — realan mismatch na trailing-space/case). Ostatak je observability/typed-result higijena (tačno, ali P2/P3 na single-user desktopu). 122.9 je najjeftiniji fix (S) sa direktnim efektom na korektnost.

---

### FM-0111 — `frmSledljivost.frm`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| GZ | Forma ne razlikuje nema-podataka / greška / nepotpun-trace (sva tri → prazna lista/„Nema podataka") | Zaključak | **Tačno** | P1 | Typed trace rezultat + eksplicitni statusi | M |
| P1-a | Nepotpun trace prikazan kao kompletan — `cmbZbirna_Change` (L462) puni listu bez upozorenja | **P1** | **Tačno** | P1 | Prikazati „NEPOTPUN TRACE" kad servis vrati `IsComplete=False` | M |
| P1-b | `Empty` = i no-data i error — `cmbZbirna_Change` tiho izlazi (L463), `PrintTracePDF` kaže „Nema podataka" (L499-500) i za grešku | P1 | **Tačno** | P1 | Razlikovati po typed rezultatu servisa | S |
| P1-c | Status „Povezano X od Y" (L289 = total − unlinked) tretira svaki neprazan `OtpremnicaID` kao validan link | P1 | **Tačno** | P2 | Računati po validnim linkovima (postoji/nije storniran/ista zbirna) | M |
| P1-d | Ručni kandidati preširoki — filtrira samo stanica+datum (L380-408), ne vozač/klasa/vrsta/zbirna/preostala kol. | P1 | **Tačno** (upis ipak ide kroz `ReassignOtkupToOtpremnica_TX`, L439 — bez korupcije) | P2 | Suziti kandidate canonical match pravilima | S |
| P1-e | `GetColumnIndex` u kritičnim tokovima (`LoadZbirne` L258, candidate L370-376, PDF L510-514/540-541) | P1 | **Tačno** | P2 | `RequireColumnIndex` za obavezne kolone | S |
| P1-f | PDF nepotpun bez oznake; prijem kg po direktnom `BrojZbirne` (L543-544) vs trace po `OtpremnicaID` lancu → različita semantika u istom dokumentu | P1 | **Tačno** | P1 | Ista veza za summary i detalje; blok/oznaka kad je trace nepotpun | M |
| P2-a | Zbirna po poslovnom broju bez stabilnog ID-a (`cmbZbirna` samo `BrojZbirne`, dedup L264-271) | P2 | **Tačno** | P2 | Hidden `ZbirnaID` + prikazni broj | M |
| P2-b | Normalizacija broja neujednačena (direktna string poređenja L518, L544) | P2 | **Tačno** | P2 | Ista normalizacija kao auto-link | S |
| P2-c | Lookup praznine nevidljive (stanica/vozač/kooperant/kupac → prazan naziv, L468/L524-528) | P2 | **Tačno** | P3 | Orphan oznaka | S |
| P2-d | `UserForm_Initialize` (L92-97) bez lokalnog `On Error GoTo EH` | P2 | **Tačno** (`Activate` ima EH L27, `Initialize` nema) | P2 | Dodati kontrolisani init EH | S |
| P2-e | Designer caption `UserForm1` (L3) + zastareli header `frmOtkupniBlokovi` (L18) | P2 | **Tačno** | P3 | Ispraviti caption/komentar; lokalizovati hardcode | S |
| Poz | Ručno povezivanje kroz TX servis (L439), hidden ID kolone (L230/237/244), `ExcludeStornirano`, odmah prikaz nepovezanih (L95), mouse-wheel detach (L584-603), refresh posle akcija | Pozitivno | **Kontekst-Pozitivno** | — | Zadržati | — |

Bilans FM-0111: Svi nalazi tačni. Najveći realni rizik = **lažni utisak kompletnog traga** (P1-a/P1-f): štampani PDF meša dve semantike povezivanja i ne obeležava nepotpunost. Nema direktne korupcije (ručni upis ide kroz proveren TX). P1-d je precenjeno kao P1 — TX validacija štiti bazu, ostaje samo UX rizik pogrešnog izbora → P2. Većina fixeva zavisi od typed trace rezultata iz FM-0105/122.3-4 (zajednički koren).

---

### FM-0112 — `frmAgrohemija.frm`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 3 | Izlazna cena u korpi ≠ knjižena cena — korpa snapshot-uje `cena/vrednost` (L493-494, L551-552), a `btnZavrsiIzlaz` zove `SaveMagacin` BEZ `overrideCena` (L623-630) → `SaveMagacin` ponovo čita cenu; **ulaz ISPRAVNO prosleđuje cenu (L843)** | **P1** | **Tačno** (asimetrija ulaz/izlaz je dokaz da je oversight) | **P1** | Proslediti `m_KorpaIzlaz(i).cena` kao `overrideCena`: `SaveMagacin(Date, art, MAG_IZLAZ, kol, koopID, parcID, brDok, "", "", m_KorpaIzlaz(i).cena)` | S |
| 4 | Svi dokumenti sa `Date` — izlaz (L624), ulaz (L834), početni dug (L180) bezuslovno `Date`; nema polja za datum | P1 | **Tačno** | P2 | Polje datuma (default `Date`) uz validaciju/permisiju za backdating | M |
| 5 | Legacy display-string identiteti („Ime Prezime (ID)" L227-228; „Naziv [JM] (ID)" L247; `ExtractIDFromDisplay`) | P1 | **Tačno** | P2 | Dvokolonski combo sa hidden stabilnim ID-em (`modComboBinding`) | M |
| 6 | Loaderi ne filtriraju aktivne + `If CStr(data(i,1))<>""` (L226/246/697) gejtuje po fizičkoj koloni 1 iako su named indeksi izračunati | P1 | **Delimično** (col1-gejt fragilnost=Tačno; „aktivni filter" zavisi da `Aktivan` postoji u šemi — **Nije proverivo statički** za sve master tabele) | P2 | Gejtovati po `colID`; aktivni filter samo gde kolona postoji | S |
| 7 | `GetColumnIndex` za obavezne UI kolone (L220-222, L240-242, L689-693) | P1 | **Tačno** (domain read-modeli ipak koriste `RequireColumnIndex`) | P2 | `RequireColumnIndex` u loaderima | S |
| 8 | Početni dug: `InputBox`+`IsNumeric`+`CDbl` (L156-167), locale-zavisno | P1 | **Tačno** | P2 | Centralni strogi parser + prikaz parsirane vrednosti pre potvrde | S |
| 9 | Semantika količine: izlaz=broj pakovanja→kg (L510-512), ulaz=ukupna kol.×cena (L741, L769) | P1 | **Tačno** | P1 | Eksplicitno označiti oba modela u UI + konverzija; sprečiti da operater unese pakovanja pri prijemu | M |
| 10 | Stock check po trenutku, korpa nema reserve model (add L515-533, finish `ValidateKorpaIzlazStanje` L612/933) | P1 | **Tačno** (re-check pri završetku postoji) | P2 | Pri rollback-u prikazati problematičan artikal, zadržati korpu | M |
| 11 | `SaveMagacin` guta typed grešku, forma diže generičku 4301 (L632-637) | P1 | **Tačno** (vidi FM-0113) | P1 | Propustiti typed razlog iz `SaveMagacin` | M |
| 12 | Korpa bez uređivanja/brisanja pojedinačne stavke (samo `ClearKorpaIzlaz`/`Ulaz` L673-677/885-889) | P2 | **Tačno** | P2 | „Ukloni stavku" dugme (indeks u UDT nizu) | M |
| 13 | Forma meša oblasti (ulaz/izlaz/preporuka/dug/početni dug/KPI) | P2 | **Dizajnersko ograničenje** | P3 | Zadržati orchestration, write ostaje u servisima (već jeste) | L |
| 14 | Designer caption `UserForm1` (L3) + delom hardcoded tekstovi | P2 | **Delimično** (caption=Tačno; mnogi tekstovi VEĆ idu kroz `Poruka()` — hardcode je manjinski, npr. „+ Dodaj u korpu", „Izaberite artikal!") | P3 | Ispraviti caption; lokalizovati preostale literale | S |
| 2 | Pozitivno: typed UDT korpe, outer TX nad `tblMagacin`, agregatna provera pre izlaza, pozitivan ceo broj pakovanja, `Pakovanje>0` invariant, override cena za ulaz, guard za dupli „Početni dug" (L116), mouse-wheel detach | Pozitivno | **Kontekst-Pozitivno** | — | Zadržati | — |

Bilans FM-0112: Nalaz #3 (price snapshot vs reprice) je **najvredniji nalaz u celom setu** — realna finansijska/audit nekonzistentnost, potvrđena asimetrijom sa ulaznim tokom (ulaz prosleđuje cenu, izlaz ne), sa jasnim minimal-delta fixom (jedan poziv, napor S). #9 (dve semantike količine) je drugi realni rizik (pogrešan prijem umanji stanje za faktor pakovanja). #6 i #14 su **Delimično** (aktivni-filter zavisi od šeme; većina tekstova je već lokalizovana). Ostalo tačno ali P2/P3.

---

### FM-0113 — `modAgrohemija.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | `ValidateMagacinInput` (L593-616) proverava samo sintaksu (artikalID/tip/kolicina/koop-za-izlaz); bez postoji/aktivan artikal, koop, dobavljač, parcela↔koop, datum, cena, dupli dok | **P1** | **Tačno** | P2 | Referencijalne provere pre upisa (postojanje+aktivnost preko `FindRows`/`LookupValue`) | M |
| 2 | Ulaz ne zahteva dobavljača — `dobavljacID` u signaturi ali neproveren; prazan → `(Nepoznat)` u izveštaju (L536) | P1 | **Tačno** | P2 | Odlučiti: obavezan dobavljač za realni ulaz, ili tipovi ulaza bez dobavljača sa reason code-om | S |
| 3 | Cena 0 tiho — `overrideCena`/master nenumerički → `cena=0`, red upisan `Cena=0,Vrednost=0` (L109-120, L130) | P1 | **Tačno** | **P1** | Za realne artikle zahtevati validnu cenu > 0 (osim rezervisanog `ART_POCETNI_DUG`) | S |
| 4 | `SaveMagacin` guta originalnu grešku (L145-149 log+`Debug.Print`+`""`); `_TX` diže generičku `4210` (L179) → `4205 Nedovoljno stanje` → „SaveMagacin nije uspeo" | P1 | **Tačno** | P1 | Core re-raise typed greške; prevod u poruku samo na UI boundary; `_TX` očuva broj/source | M |
| 5 | Početni dug parcijalni side-effect — `EnsureArtikalPocetniDug` PRE TX (L277), snapshot samo `tblMagacin` (L282) → fail knjiženja ostavlja virtuelni artikal | P1 | **Tačno** (FM sam priznaje „nije nužno korupcija"; `Ensure` idempotentan L325, artikal se reuse-uje) | P3 | Setup migracija unapred kreira artikal, ili snapshot i `tblArtikli` | S |
| 6 | Ensure guta grešku (L345-346) a `BookPocetniDug` nastavlja; `Validate` ne proverava artikal → moguć `MAG` red bez master artikla | P1 | **Tačno** | P2 | Exact-one post-condition provera rezervisanog artikla u `BookPocetniDug` | S |
| 7 | Nema idempotency/duplicate zaštite — uvek nov `MAG-*` (L100) | P1 | **Delimično** (Tačno za servis; **korpa flow je delom mitigovan** — posle uspeha `ClearKorpaIzlaz` L646, pa dupli klik nailazi na praznu korpu; ostaje retry/programski poziv) | P2 | `DocumentID/OperationID` za multi-row korpu | M |
| 8 | Stock validation ne razlikuje data-quality od nule — `GetArtikalStanje` (L618-633) vraća 0 za više uzroka | P1 | **Tačno** | P3 | Typed „nepoznat artikal" vs „stanje 0" | S |
| 9 | `GetParceleByKooperant` ne filtrira aktivne parcele (L4-65); invalid površina→0 tiho (L46-50) | P1 | **Delimično** (nema active-filter=Tačno; da li `Parcele` ima „aktivan" status — **Nije proverivo statički**) | P2 | Aktivni filter ako kolona postoji; površina→warning umesto tihe 0 | S |
| 10 | `CalculatePreporuka` vraća 0 za više grešaka (L67-78: ne postoji/doza fali/nenumerička/lookup fail/stvarna 0) | P1 | **Tačno** | P3 | Bool+output+reason ili typed rezultat | S |
| 11 | `GetMagacinStanje` sabira samo `MAG_ULAZ/IZLAZ` (L392-396); nepoznat tip nevidljiv bez warninga | P2 | **Tačno** | P3 | Health-check za invalid tipove | S |
| 12 | Lookup master → prazni nazivi bez markera (L414-416, L488-489) | P2 | **Tačno** | P3 | Orphan marker | S |
| 13 | Typo alias — `ReportStanjePoDobavljacu` (L504) prosleđuje `ReportStanjePoDoabvljacu` (L508); **canonical je pogrešno napisan, tačan je forwarder** | P2 | **Tačno** | P3 | Preimenovati canonical u ispravan naziv, typo označiti deprecated | S |
| 14 | `SaveMagacin_TX` `userId="Operator"` (L193, L236) | P2 | **Tačno** | P3 | Stvarni auth user | S |
| Poz | `RequireColumnIndex` u read-modelima, `ExcludeStornirano` (L363/433/515/572), izuzimanje `ART_POCETNI_DUG` iz stanja (L387), `_TX` rollback+monitoring, reverzibilan početni dug, stock u domain sloju (L122-127), override cena za ulaz | Pozitivno | **Kontekst-Pozitivno** | — | Zadržati | — |

Bilans FM-0113: Nalazi većinom tačni. Dva su suštinska: **#3** (tiha cena 0 potcenjuje dug/vrednost — finansijski, fix S) i **#4** (gubitak typed greške — direktno objašnjava slabu dijagnostiku iz FM-0112 #11). #7 i #9 su **Delimično** (korpa dupli-klik je mitigovan clear-om; parcela active-status nije statički proveriv). #5 (parcijalni bootstrap) je realno bezopasan jer je `Ensure` idempotentan → P3. Napomena o KI-006: `GetMagacinStanje` KOREKTNO izuzima `ART_POCETNI_DUG` (L387) — registrovani KI-006 se odnosi na drugi export (`ExportMagacinKoop` u PWA sync sloju), koji NIJE u ovom modulu; ovde nema regresije tog tipa.

---

## Sažetak (cross-FM)

Sve FM stavke su verifikovane protiv koda; nijedan nalaz nije **Netačno**. Distribucija: pretežno **Tačno**, sa nekoliko **Delimično** (FM-0112 #6/#14, FM-0113 #7/#9 — zbog schema-drift neproverivosti aktivnog statusa i already-mitigovanih tokova) i **Dizajnersko ograničenje** (FM-0096 116.19 SWMR race — van single-writer modela; 116.30-34 predlozi).

Rangiranje po stvarnoj hitnosti (P1), sa minimal-delta fiksom:
1. **FM-0112 #3** — izlazna cena: proslediti `m_KorpaIzlaz(i).cena` kao `overrideCena` (napor **S**). Najjači nalaz — finansijska nekonzistentnost, potvrđena asimetrijom sa ulazom.
2. **FM-0113 #3** — tiha cena 0: zahtevati validnu cenu za realne artikle (**S**).
3. **FM-0113 #4 / FM-0112 #11** — očuvati typed grešku iz `SaveMagacin` (**M**).
4. **FM-0105 122.3 / FM-0111 P1-a,f** — nepotpun trace prikazan/štampan kao kompletan; direktan `BrojZbirne` prolaz + `IsComplete` (**M**).
5. **FM-0105 122.9** — normalizacija broja zbirne u trace-u (**S**).
6. **FM-0096 116.6+116.20+116.21** — canonical mirror identitet + provera rezultata pre `VozacID=StanicaID` (**M**).

Glavna korekcija FM kalibracije: **P0 oznake u FM-0096 su precenjene za single-writer desktop** sa postojećim mitigacijama (backfill/hook/auto-hladnjaca kreiraju mirror); realno P1. Većina „P1 typed-result/observability" nalaza su engineering-quality (P2/P3), ne aktivni bug-ovi.

---

## v142 blok K5 — test moduli: modTestStorno, modTestStornoCentar, modTestPalete, modIzvestajTests, stub sweep, modSEFTests (FM-0097/0098/0099/0100/0101/0136) [sidro origin/main v2.24.0]

Sve infrastrukturne tvrdnje verifikovane. Sažetak potvrda pre tabela:

- **`clsTransaction`** (`clsTransaction.cls:41,51,108-147`): `AddTableSnapshot` snima celu `DataBodyRange.Value2`; `RestoreTable` briše `DataBodyRange` i prepisuje celu matricu — whole-table restore, NIJE selektivan po `SVT-`/`TST-` vlasništvu. Mehanizam SWMR data-loss-a je realan, ali app je single-writer desktop.
- **`StornoSelectedBlocks_TX`** (`modStornoFlow.bas:1395,1406,1428`): interni `_TX` snima i commit-uje `TBL_STORNO_ZURNAL` + `BeginStornoOp` journalizuje. Potvrđuje da outer test koji NE snima zurnal ostavlja žurnal-residue.
- **`modE2EReleaseGate.bas:23-49`**: gate zove SAMO `RunGoogleSyncSmokeSuite/RunMasterSyncSmokeSuite/RunNovacSmokeSuite/RunFakturaSmokeSuite/RunBusinessFlowProSuite/RunProductionHealthCheck` (+ manual GAS). NIJEDAN od auditovanih 5 modula nije u gate-u → „E2E false-green" nalazi su LATENTNI/uslovni (AUD-039).
- **`modRelease.bas:5,18`**: „objavi src-vba kod" iz `SRC_FOLDER=...\src-vba\` → svi `.bas` (uklj. test module) idu klijentima kroz self-update. Potvrđuje shipped-destructive-macros rizik (AUD-039).

---

### FM-0097 — `modTestStorno.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | Stale doc „8 scenarija" a runner ima 22 (118.3/118.26) | — | **Tačno** — header komentar `:16` kaže 8; `:76-97` poziva T01–T22 | P3 | Osveži header/komentar na 22 + scenario manifest | S |
| 2 | `TBL_STORNO_ZURNAL` van outer snapshot-a (118.6/118.7) | P0 | **Tačno** — snapshot lista `:58-68` bez zurnala; `_TX` sloj demonstrira journal (`modStornoFlow.bas:1406`). Tačan broj revers-žurnal redova po T06/T07/T12/T13 = **Nije proverivo statički** bez čitanja revers grana | P2 | Dodaj `tx.AddTableSnapshot TBL_STORNO_ZURNAL` (1 linija) | S |
| 3 | Whole-table rollback nije SWMR-safe (118.8) | P0 | **Tačno (mehanizam)** — `clsTransaction.cls:108-147`; ali **Dizajnersko ograničenje** za single-writer desktop | Prihvaćeno | Dokumentuj single-writer pretpostavku u headeru; ne refaktoriši u disposable-copy | S |
| 4 | Outer rollback + inner commit nije prava nested TX (118.9) | — | **Tačno** — `_TX` prave sopstvene `clsTransaction`+`CommitTx`; opis je tačan, posledica = red 2 | P3 | Bez izmene koda; napomena u dokumentaciji | S |
| 5 | Deferred AutoSave ostaje zakazan (118.10) | — | **Nije proverivo statički** — zavisi od `MarkDirtyAndSchedule` u inner `CommitTx` (nije čitano); plauzibilno tačno | P3 | Suspend/restore AutoSave oko run-a ako se potvrdi | M |
| 6 | CSV crash journal kontaminiran (118.11) | — | **Delimično** — fixture `SvAppend` (`:752-763`) koristi direktan `ListRows.Add`, NE `AppendRow` → fixture ne piše CSV; residue moguć samo iz inner `_TX` (nije proverivo statički) | P3 | Test journal sink ili offset restore — samo ako se potvrdi | M |
| 7 | Schema ensure pre korisničke potvrde (118.12) | — | **Tačno** — `EnsureStornoVezeSchemaCore` `:42` pre MsgBox `:49` | P3 | Pomeri ensure posle potvrde | S |
| 8 | Nema `tblStornoZurnal` precondition (118.13) | — | **Tačno** — samo `EnsureStornoVezeSchemaCore`; revers testovi zavise od zurnal šeme | P3 | Dodaj `EnsureStornoZurnalSchemaCore` u preflight | S |
| 9 | Interruption ostavlja fiksne `SVT-*`, nema residue preflight (118.14) | — | **Tačno** — fiksni ID-jevi; NEMA residue provere (za razliku od `modTestPalete:67-73`). Elevirano jer je modul shipped (AUD-039) | P1 | Preslikaj residue-preflight iz `modTestPalete` (abort ako `SVT-*` postoji) | S/M |
| 10 | `SvAppend` tiho preskače missing kolone (118.15) | — | **Tačno** — `:761 If ci > 0 Then` (fail-soft, klasa false-green) | P2 | `RequireColumnIndex` za obavezna fixture polja | S/M |
| 11 | Fixture bypassuje canonical save (118.16) | — | **Dizajnersko ograničenje** — namerno za izolovan engine test; **Kontekst-Pozitivno** | P3 | Reklasifikuj kao „engine integration", ne E2E | S |
| 12 | Nema exact-one identity zaštite (118.17) | — | **Tačno** — nema `CountRowsByKey` pre/posle seed-a; preklapa se s #9 | P3 | Objedini sa residue-preflight iz #9 | S |
| 13 | T14 osetljiv na concurrent pending (118.18) | — | **Delimično** — `:446-456` koristi relativni `before+2` (robusno za single-writer); concurrent-writer je teorijski | P3 | Bez izmene (relativni count je dovoljan single-writer) | S |
| 14 | Fatal error provenance nije sačuvan (118.19) | — | **Tačno** — EH `:104-109` radi `On Error Resume Next`→rollback→koristi `Err.description` (nije capture-ovan pre rollback-a; `modTestPalete:112` to radi bolje) | P2 | Sačuvaj `Err.*` u lokale pre rollback-a | S |
| 15 | Rollback uspeh se ne verifikuje (118.20) | — | **Tačno** — nema post-cleanup hash/row-count provere | P3 | Pre/post fingerprint (nice-to-have) | M |
| 16 | PASS/FAIL nije machine-readable / rizik u E2E (118.21) | — | **Tačno** — `Public Sub`, MsgBox/Debug.Print, bez Boolean-a; **potvrđeno NIJE u E2E gate-u** → false-green je uslovan (AUD-039) | P3 | `...Core() As Result` + `...Interactive()` split | M |
| 17 | Interaktivni modal sprečava automatizaciju (118.22) | — | **Tačno** — Yes/No `:49` + modalni summary `:885` | P3 | Isto kao #16 (Core/Interactive) | M |
| 18 | Monitoring/log residue nije u rezultatu (118.23) | — | **Nije proverivo statički** — zavisi od inner monitoring emitovanja; plauzibilno | P3 | TEST marker/sink ako se potvrdi | M |
| 19 | Brojači mere assertions, ne scenarije (118.24) | — | **Tačno** — `Chk/ChkEq/ChkEqD` inkrementiraju `mPass/mFail` po assertion-u | P3 | Dodaj case-level `Started/Passed/Failed` | M |
| 20 | Nema timing/build/env identiteta (118.25) | — | **Tačno** — report `:870-886` bez RunID/SHA/timestamp | P3 | Dodaj metapodatke u report | S |
| 21 | Hardening/arhitektura/ocena (118.27–118.32) | — | **Predlozi (ne-nalazi)** — validne smernice; scenario-spec kvalitet visok (sačuvati svih 22) | P3 | Prihvatiti selektivno; ne prepisivati scenario logiku | L |

**Bilans FM-0097:** Skoro sve tvrdnje **tačne** protiv koda. Dva FM „P0" (SWMR whole-table, 118.8) su realni po mehanizmu ali **Prihvaćeno** za single-writer. Najmaterijalnije i najjeftinije: **zurnal-snapshot gap (#2, 1 linija)** i **residue-preflight (#9)** — potonji je instanca AUD-039 (shipped destruktivni makro bez čvrstog guard-a, ublažen samo Yes/No potvrdom). Duplira se sa AUD-016 (modul je jedan od test dubleta). Fail-soft `SvAppend` (#10) je klasičan false-green izvor.

---

### FM-0098 — `modTestStornoCentar.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | „22 pozvane test procedure" (120.1) | — | **Netačno (korekcija FM)** — `Test_StornoCentar_All :21-43` poziva **23** procedure (dodat `Test_ImpactHeaderSum_Auto :43`). FM podbrojao za 1 | P3 | Ispravi FM broj na 23; header uskladi | S |
| 2 | `Test_StornoSelectedBlocks_Auto` ostavlja žurnal (120.5) | P0 | **Tačno** — `:640-642` snima Otkup/Ambalaza/Novac BEZ `TBL_STORNO_ZURNAL`; `StornoSelectedBlocks_TX` interno journalizuje i commit-uje (`modStornoFlow.bas:1406`). Kontrast: `Test_StornoJournalPartialClass_Auto :190` snima zurnal | P2 | Dodaj `tx.AddTableSnapshot TBL_STORNO_ZURNAL` u taj test (1 linija) | S |
| 3 | Posledice mrtvih journal ops (120.6) | — | **Tačno** — residue vidljiv u Storno centru/undo/recovery; opovrgava header „fixture NE ostaje" `:9-10` | P2 | Reši kroz #2 | S |
| 4 | Whole-table rollback nije SWMR-safe (120.7) | P0 | **Tačno (mehanizam)**, **Dizajnersko ograničenje** za single-writer (kraći exposure — 1 snapshot po testu) | Prihvaćeno | Dokumentuj; bez refaktora | S |
| 5 | Nema environment safety gate-a (120.9) | P0 | **Tačno** — `:19-45` pokreće svih 23 odmah; NEMA potvrde/TEST_MODE/workbook-provere (gore od `modTestStorno` koji bar ima MsgBox). Shipped modul (AUD-039) | **P1** | Dodaj guard: `TEST_MODE`/workbook-name/sandbox check + potvrda pre run-a | M |
| 6 | Public macro surface preširok (120.10) | — | **Tačno** — svih 23 su `Public Sub` (Alt+F8 izlaže najdestruktivniji test direktno) | P1 | `Option Private Module` + jedan kontrolisani entry; ali guard (#5) je primaran | S |
| 7 | Runner false-green + nije u E2E (120.11) | — | **Tačno** — `:44` Debug.Print, bez Boolean/raise; **potvrđeno NIJE u E2E** → uslovno (AUD-039) | P3 | Structured result pre bilo kakve E2E integracije | M |
| 8 | Assertions se broje, ne case-ovi (120.12) | — | **Tačno** — `TcChk :773` inkrementira `mPass/mFail` | P3 | Case-level metrika | M |
| 9 | Fatal error handling gubi original (120.13) | — | **Tačno** — svaki EH (npr. `:65-67`) rollback pa `Err.description`, bez capture pre; runner nema sopstveni EH | P2 | Capture `Err.*` pre rollback-a; odvoji `cleanupFailed` | S |
| 10 | Fiksni fixture + interruption residue, nema preflight (120.14) | — | **Tačno** — fiksni `SVT-*/SOP-MIX/ZUR-*`; nema residue provere | P2 | Residue-preflight po namespace-u | S/M |
| 11 | `TcSeedRow` nije fail-fast (120.15) | — | **Tačno** — `:762-771 If lo Is Nothing Then Exit Sub` + `If ci>0` (silent skip, fail-soft) | P2 | Fail-fast na missing tabelu/kolonu | S/M |
| 12 | Negativni testovi mogu proći iz pogrešnog razloga (120.16) | — | **Tačno** — testovi proveravaju samo `= False` (npr. `:344,:480`); ne razlikuju business-guard od missing table/exception | P2 | Proveri tačan reason/error-code | M |
| 13 | Direktan `ListRows.Add` zaobilazi DataAccess (120.17) | — | **Tačno** (`:765`); **Dizajnersko** za fixture, ali nema cache-invalidate | P3 | Invalidiraj cache + readback potvrda | S |
| 14 | Conditional stale-cache rizik (120.18) | — | **Nije proverivo statički** — zavisi od `BeginTableCache` stanja u run-time; teorijski tačno | P3 | Zahtevaj cache depth 0 / reset pre scenarija | S |
| 15 | Hidden schema prerequisite `EnsureRuntimeSchema` (120.19) | — | **Tačno** — header `:12-13` traži ručno; runner ne proverava; journal testovi sami zovu `EnsureStornoZurnalSchemaCore :75` → nedosledno | P3 | Jedinstven preflight svih šema | S |
| 16 | CSV crash journal nije rollbackovan (120.20) | — | **Delimično / Nije proverivo statički** — `TcSeedRow` direktan (bez CSV); inner `_TX` može pisati. Isto kao FM-0097 #6 | P3 | Test sink ako se potvrdi | M |
| 17 | Deferred AutoSave nije izolovan (120.21) | — | **Nije proverivo statički** — zavisi od inner `CommitTx→MarkDirtyAndSchedule` | P3 | Suspend/restore ako se potvrdi | M |
| 18 | Monitoring side-effecti trajni (120.22) | — | **Nije proverivo statički** — `Test_ZbirnaRecalcInPlace_Auto :446` proverava audit ali rollbackuje samo `tblZbirna`; monitoring nije snapshot-ovan (plauzibilno tačno) | P3 | Injectabilni test sink | M |
| 19 | Spec kvalitet: reused-broj/dual-class/drift/partial/dead-parent (120.23–120.31) | — | **Kontekst-Pozitivno** — potvrđeno u kodu (`:72-127,:301-326`), visoka regres-vrednost; sačuvati | — | Zadrži scenario logiku 1:1 | — |
| 20 | Test order/izolacija (120.32/120.33) | — | **Tačno** — per-test outer TX (dobro), ali dele counters/workbook/namespace; nema ambient/reentrancy preflight (`StornoOpActive`) | P3 | Ambient-state preflight | M |
| 21 | Harness contract/arhitektura (120.34–120.42) | — | **Predlozi (ne-nalazi)** | P3 | Selektivno | L |

**Bilans FM-0098:** Tvrdnje uglavnom **tačne**; jedina činjenična greška je u samom FM-u (23, ne 22 procedure — #1). Najveći realni rizik je **#5 (nula env guard-a nad shipped destruktivnim runnerom)** → **P1**, direktna instanca AUD-039 i najoštriji od svih 5 modula (nema ni MsgBox). **#2 (zurnal-snapshot gap)** je jednolinijski correctness fix koji preživljava i single-writer. SWMR P0 = **Prihvaćeno**. Preklapanje: AUD-016 (dubl sa `modTestStorno`).

---

### FM-0099 — `modTestPalete.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | Header „T01–T10" a runner zove i T11 (122.1/122.12) | — | **Tačno** — header `:29-40` broji 10; `:104` poziva `T11_RelabelBlokNaDeljenojPaleti` | P3 | Uskladi header + manifest | S |
| 2 | Whole-table rollback nije SWMR-safe (122.5) | P0 | **Tačno (mehanizam)**; komentar `:12-16` eksplicitno tvrdi „sme na živoj svesci" — **Dizajnersko ograničenje** za single-writer, ali komentar je precenjeno obećanje | Prihvaćeno (rizik) / P3 (komentar) | Ukloni „sme na živoj shared svesci" tvrdnju iz komentara | S |
| 3 | „Ne ostavlja podatke" netačno — `_TestPalete` sheet (122.6) | P0 | **Tačno** — `ReportResults :648-667` posle rollback-a kreira/`Cells.Clear`/upisuje sheet; nije u snapshot-u | P2 | Nazovi ga eksplicitnim durable artifact-om; uskladi MsgBox „sve vraćeno" `:671-678` | S |
| 4 | Spoljne side-effecte TX ne vraća (122.7) | P0/P1 | **Delimično / Nije proverivo statički** — CSV/monitoring/autosave zavise od inner `_TX`; `TstAppend :490` koristi `AppendRow` (za razliku od `SvAppend`), pa fixture jeste u CSV kanalu | P3 | TEST marker/sink; potvrdi inner ponašanje | M |
| 5 | AutoSave pretpostavka nepotpuna (122.8) | — | **Nije proverivo statički** — `CancelAutoSaveTimer` se ne zove (grep-potvrdivo da nema poziva); efekat zavisi od timer stanja | P3 | Suspend/restore AutoSave | M |
| 6 | Residue preflight parcijalan (122.9) | — | **Tačno** — `:67-68` proverava samo `TST-VOCE` i `TST-P1`, ne sve `TST-*` identitete | P2 | Proširi preflight na sve fixture prefikse | S |
| 7 | Fixture identiteti nisu run-scoped (122.10) | — | **Tačno** — fiksni `TSTPRJ-*/TST-P*/TST-Z*` | P3 | `TST-<RunID>-...` + PK registry | M |
| 8 | Scenario pokrivenost T01–T11 (122.11) | — | **Kontekst-Pozitivno** — potvrđeno (`:127-460`), jaka engine matrica; sačuvati | — | Zadrži | — |
| 9 | Assertion model (case vs assertion) (122.13) | — | **Tačno** — `mPass/mFail` broje provere; `ReportResults(clean)` razlikuje samo normal/prekid | P3 | Per-case status | M |
| 10 | Fatal error obrada (122.14) | — | **Tačno (delimično dobra)** — `:112` čuva `eDesc` pre `On Error Resume Next` (bolje od FM-0097), ali ne `Err.Number/Source`; rollback failure progutan; MsgBox tvrdi „vraćeno" i kad možda nije | P2 | Razdvoji `TEST_FAILED` vs `CLEANUP_FAILED` | S/M |
| 11 | Report artifact rizici — `Cells.Clear` bez test-owned provere (122.15) | — | **Tačno** — `:656 ws.Cells.Clear` na sheet-u imena `_TestPalete` bez provere vlasništva; ceo report pod `On Error Resume Next :643` | P3 | Proveri da je sheet test-owned pre Clear | S |
| 12 | Test sloj zaobilazi save API (122.16) | — | **Dizajnersko ograničenje / Kontekst-Pozitivno** — namerno; header priznaje `:29` | P3 | Reklasifikuj kao paleta-engine integration | S |
| 13 | Master-data rollback SWMR (122.17) | — | **Tačno (mehanizam)** — `SeedMasterData :466` menja globalne `tblKulture/tblTipAmbalaze`; **Dizajnersko** za single-writer | Prihvaćeno | Rezervisan test-master ili disposable | M |
| 14 | Module-level state `mSkipPaletize` ne vraća (122.18) | — | **Tačno** — `:64 SetPaletizeSkip False` na startu, nema restore prethodne vrednosti; nema javnog getter-a | P2 | Očitaj/vrati prethodno stanje u cleanup-u (treba getter) | S/M |
| 15 | Runtime model/hardening (122.19–122.23) | — | **Predlozi (ne-nalazi)** | P3 | Selektivno | L |

**Bilans FM-0099:** Modul je najbolje-ograđen mutacioni suite od tri storno/paleta (paletiranje-enabled + delimičan residue-check + Yes/No potvrda). Tvrdnje **tačne**; SWMR P0 = **Prihvaćeno/Dizajnersko** (single-writer), ali **komentar `:12-16` je objektivno precenjeno obećanje** i vredi ga omekšati. Konkretni jeftini fixevi: `_TestPalete` „nije-promena" laž (#3), parcijalan preflight (#6), `mSkipPaletize` restore (#14). Nije destruktivan nad poslovnim podacima kao storno moduli (rollback + engine-only).

---

### FM-0100 — `modIzvestajTests.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | Naziv modula: komentar `modSmokeTestIzvestaj` vs `VB_Name modIzvestajTests` (123.1) | — | **Tačno** — `:1` vs `:5` | P3 | Uskladi komentar sa `VB_Name` | S |
| 2 | `EMPTY`/`NOT ARRAY`/`INVALID ARRAY` nisu failure (123.4) | P0 | **Tačno** — `Smoke_ArrayShape :200-217` vraća tekst; `Smoke_RunReport :185` samo ispisuje; runner bezuslovno `"...OK" :172`. Read-only modul → nema data rizika, ali čist false-green | P2 | Tretiraj te statuse kao FAIL osim gde je EMPTY eksplicitno očekivan; brojač FAIL | M |
| 3 | Per-report EH ne hvata greške producer-a — arg eval order (123.5) | P0 | **Tačno** — `Smoke_RunReport "ReportSaldoOM", ReportSaldoOM(...)` `:50`: VBA evaluira argument PRE ulaska, pa exception ide u suite-level EH `:176`, ne u `Smoke_RunReport.EH :191`. Realna harness greška (prvi izuzetak prekida ceo suite) | P2 | Pozovi svaki report u zasebnom case-wrapper-u koji stvarno hvata producer exception | M |
| 4 | Fatalna greška se guta (123.6) | P0 | **Tačno** — suite EH `:176-183` ispisuje `Err.*` ali bez `Err.Raise`/Boolean/artifact; **NIJE u E2E** → uslovni false-green (AUD-039) | P2 | Vrati `Succeeded` Boolean/result | S/M |
| 5 | Sample selection nedeterministički (123.7) | — | **Tačno** — `Smoke_FirstValue :219` uzima fizički prvi red bez provere aktivnosti/perioda | P3 | Biraj aktivan entitet sa podacima u periodu | M |
| 6 | Period zavisi od dana pokretanja (123.8) | — | **Tačno** — `:21-22` `DateSerial(Year(Date),1,1)`..`Date` | P3 | Fiksni/data-driven stabilan period | S |
| 7 | Nema contract provere po reportu (123.9) | — | **Tačno** — proverava se samo oblik `rows x cols`, ne broj kolona/total/tipovi | P3 | Očekivani br. kolona + min redova po reportu | M |
| 8 | SKIP nema suite-level posledicu (123.10) | — | **Tačno** — `Smoke_Skip :196` samo ispisuje; nema coverage threshold | P3 | Brojač SKIP + prag obaveznih | S |
| 9 | Helperi gutaju schema/data greške (123.11) | — | **Tačno** — `Smoke_FirstValue` EH `:241-242` vraća `""` → schema regres izgleda kao „Nema ID" (fail-soft, false-green klasa) | P2 | Razlikuj `NoData` od schema/data-access FAIL | S/M |
| 10 | Ne validira poslovnu matematiku (123.12) | — | **Tačno** — nema poznatih numeričkih scenarija (saldo/ponder/manjak) | P3 | Deterministički fixture regres u izolaciji | L |
| 11 | Read-only, širok API, brz compat smoke (123.13) | — | **Kontekst-Pozitivno** — potvrđeno read-only (nema mutacije poslovnih tabela); najbezopasniji od 5 | — | Sačuvati kao smoke, ne kao dokaz tačnosti | — |
| 12 | Podela testova/result contract/matrix (123.14–123.18) | — | **Predlozi (ne-nalazi)** | P3 | Selektivno | M/L |

**Bilans FM-0100:** Sve **tačno**. Za razliku od storno/paleta modula, ovaj je **read-only** (nema fixture/rollback/data rizika), pa tri FM „P0" nisu data-loss već isključivo **false-green harness** greške — realne, ali niske štete jer modul nije u E2E gate-u i ne mutira podatke. Najzanimljiviji je **#3 (arg-eval-order)** — istinska VBA zamka koja obara pouzdanost celog runnera. Fail-soft `Smoke_FirstValue` (#9) je ista klasa kao `CountRows→0` iz konteksta.

---

### FM-0101 — Stub sweep: `modML` / `modKvalitet` / `modMeteo` / `modHladnjaca`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | Četiri prazna placeholder modula, bez procedura (116.1–116.5) | — | **Tačno** — svaki 7 linija (`Attribute`+`Option Explicit`+roadmap komentar+`TODO: Implementierung`); FM kaže „8 linija" (minorna netačnost) | P2 | — | — |
| 2 | Nema production callere; runtime rizik ~0 (116.6) | — | **Tačno** — bez javnih članova; ne mogu proizvesti false-green/side-effect/compile grešku | Prihvaćeno | — | — |
| 3 | Lažna arhitektonska signalizacija (naziv sugeriše da sloj postoji) (116.2–116.5) | — | **Tačno** — posebno `modHladnjaca` (postoje realni prerada/paleta/monitoring moduli) | P3 | Roadmap/ADR umesto praznih modula | S |
| 4 | Ukloniti sva četiri iz `src-vba` import scope-a (116.7/116.8) | P2 | **Tačno / opravdano** — smanjuje broj modula bez vrednosti; ali su ipak shipped kroz `modRelease` | P2 | Ukloni iz `src-vba`; ideje u backlog | S |

**Bilans FM-0101:** Potpuno **tačno** (uz sitnu FM grešku 8 vs 7 linija). Nula funkcionalnog rizika — čist repo/workbook cleanup, **P2**. Jedina veza sa ostatkom audita: i ovi prazni moduli idu klijentima kroz `modRelease` (isti pipeline kao shipped test moduli), pa uklanjanje ima i higijensku vrednost.

---

### FM-0136 — `modSEFTests.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | Suite failure ne propagira caller-u | P0 | **Tačno** — svi `Run*Suite` EH→`LogFatal`→`FinishSuite`→normal return (`:66-69`); `FinishSuite :765-786` samo MsgBox, bez Boolean/raise; **NIJE u E2E** → uslovno (AUD-039) | P2 | `SEFTestSuiteResult` sa `Succeeded`; raise/return na FAIL/FATAL | M |
| 2 | „Offline suite" ipak menja workbook | P1 | **Tačno** — `RunSEFOfflineSuite :41` zove `InitSEFTestLog :825` koji dodaje `SEF_TEST_LOG` sheet + redove | P3 | Preimenuj u „no business-table mutation" ili log u eksterni fajl | S |
| 3 | Live testovi menjaju fakture/SEF bez cleanup-a | P1 | **Tačno** — `SendInvoiceToSEF_TX/RefreshSEFStatus_TX/Cancel/Storno` nepovratni; nema test-owned registra ni cleanup-a | P2 | Dedikovane test fakture u test env; nikad proizvoljna poslovna | M |
| 4 | Public live/destructive bez auth/role guard | P1 | **Delimično** — jesu `Public Sub`, ALI postoje guardovi: `RequireLiveSEFTestsAllowed :654` (config flag) + `RequireCancelStornoTestsAllowed :1164` + `ConfirmDangerousSEFMutation :1175` (kucana potvrda). Nedostaje samo AgriX user/role provera. Najbolje ograđen modul od 5 | P2 | Dodaj admin/dev role check uz postojeće flag-ove | S |
| 5 | Production detection = string heuristika | P1 | **Delimično** — `IsLikelyProductionSEF :687` PRVO proverava `SEF_ENV=PROD/PRODUCTION` (već je enum koji FM predlaže!), tek onda substring `DEMO/TEST/SANDBOX`. Rizik realan samo ako `SEF_ENV` prazan I prod-URL sadrži „TEST" | P2 | Učini `SEF_ENV` obaveznim (fail-closed) + host allowlist; substring kao fallback | S |
| 6 | `FindFirstFakturaID` bira proizvoljnu fakturu | P1 | **Tačno** — `:626` prvi nestorniran red, bez provere statusa/kompletnosti/test-owned | P2 | Zahtevaj eksplicitan test-owned FakturaID | S |
| 7 | SKIP može sakriti regression | P1 | **Tačno** — `Test_ValidateFakturaForSEF_DoesNotCrash :332` bilo koja greška→SKIP; DTO datum→SKIP `:269`; live SKIP-ovi | P2 | SKIP samo za eksplicitno van-scope precondition; inače INCONCLUSIVE/FAIL | M |
| 8 | Negativni testovi prihvataju bilo koju grešku | P1 | **Tačno** — `Test_PayloadValidationRejectsEmpty :287` bilo koja greška=pass; `AssertTransitionBlocked :410`; `Test_GetJsonNumericIdLiteralRaises :1440` samo `Err.Number<>0` iako komentar traži `ERR_SEF_VALIDATION` | P2 | Proveri `Err.Number/Source` (tačan kod) | S/M |
| 9 | Ne pokriva false-success response parser | P1 | **Tačno (coverage-gap)** — u modulu nema testa za HTTP200 prazan body/malformed JSON/unknown status/202-bez-docID; tvrdnja o `clsSEFResponse` ponašanju = **Nije proverivo statički** ovde (drugi fajl) | P2 | Dodaj response-parser false-success matricu | M |
| 10 | Mutable snapshot/line regression nije testiran | P1 | **Tačno (coverage-gap)** — testira se samo initial build + XML substring-ovi | P3 | Dodaj mutation/final-revalidation testove | M |
| 11 | UBL validacija površna | P1 | **Tačno** — `Test_BuildDtoAndUBL :259-263` samo `InStr` substring-ovi (`<Invoice`,`<cbc:ID>`...), bez XML/schema/totals/PDV | P3 | Well-formed XML + schema + financijski invarianti | M |
| 12 | Live send prihvata REJECTED kao pass | P1 | **Tačno** — `:468-473 Case WF_SEF_REJECTED → LogPass` | P3 | Razdvoji connectivity/submission/business-accepted smoke | S |
| 13 | Refresh error posle cancel/storno se guta | P1 | **Tačno** — `:1006-1009` i `:1114-1117` `On Error Resume Next`+`Err.Clear` oko `RefreshSEFStatus_TX` | P2 | Evidentiraj refresh failure kao FAIL/typed warning | S |
| 14 | Cancel status očekivanje preširoko | P1 | **Tačno** — `:1025-1030` prihvata i `DRAFT`/`NEW` kao post-cancel | P2 | Traži jasan eksternI cancel status; inače INCONCLUSIVE | S |
| 15 | Event-log count failure izgleda kao nula | P1 | **Tačno** — `CountSEFEventsForFaktura :1216-1217` EH→0 (fail-soft, ista klasa kao kontekstni `CountRows→0`) | P2 | Razdvoji schema/missing-table grešku od „nema događaja" | S |
| 16 | Test log nema RunID/build correlation | P2 | **Tačno** — `AppendTestLog :839-858` piše Timestamp/Kind/Name/Status/Details/Operator | P3 | Dodaj SuiteRunID/build/env/FakturaID | S |
| 17 | Windows username umesto auth korisnika | P2 | **Tačno** — `:857 Environ$("Username")` | P3 | Beleži i AgriX user | S |
| 18 | `Total` = assertion/log outcome, ne test case | P2 | **Tačno** — `LogPass/LogFail/LogSkip :788-810` svako inkrementira `m_Total` | P3 | Odvoji Cases od Assertions | S |
| 19 | PATCH 5/PATCH 10 stale migracioni komentari | P2 | **Tačno** — `:1244-1256`, `:1338-1352` sadrže „dodaj na kraj modula" instrukcije | P3 | Obriši migracione komentare, zadrži svrhu/contract | S |
| 20 | Pozitivni: live/offline razdvojeni, config flag + kucana potvrda, transition matrica, idempotency, doc-ID shape | — | **Kontekst-Pozitivno** — potvrđeno (`:335-412,:1257-1420`); najsigurniji guard-model od svih 5 | — | Zadrži | — |
| 21 | Preporuke P0/P1/P2 + regres matrica | — | **Predlozi (ne-nalazi)** | P2/P3 | Selektivno | L |

**Bilans FM-0136:** Tvrdnje pretežno **tačne**; dve „P1" su **Delimično** jer FM potcenjuje postojeće guardove — modul JESTE najbolje ograđen (config flag + kucana potvrda + `SEF_ENV` enum već postoji, #4/#5), pa je najmanje pogođen AUD-039 za live/destructive grane (rezidual: offline suite bez guard-a + user-role provera). Dominira klasa **fail-soft/SKIP/any-error-pass false-green** (#7,#8,#13,#14,#15) i **coverage-gap** (#9,#10,#11). Jedini strukturni P0 (nepropagacija) je uslovan jer modul nije u E2E gate-u.

---

## Zbirni bilans (svih 6 FM unosa)

1. **Verifikacija:** ~95% FM tvrdnji **Tačno** protiv koda. Otkrivene **dve FM sopstvene netačnosti**: FM-0098 kaže 22 pozvane procedure — realno **23** (`Test_StornoCentar_All:21-43`); FM-0101 kaže 8 linija — realno **7**. Nekoliko „P0/P1" je **Delimično** (SEF guardovi postoje) ili **Nije proverivo statički** (AutoSave/CSV/monitoring — zavise od inner `_TX` grana koje nisam čitao).

2. **SWMR whole-table data-loss** (FM P0 u 0097/0098/0099): mehanizam **potvrđen** (`clsTransaction.cls:108-147`), ali **Prihvaćeno/Dizajnersko** za single-writer desktop. Preporuka: NE ulagati u disposable-copy refaktor; samo omekšati precenjeni komentar u `modTestPalete:12-16`.

3. **Najmaterijalniji, deployment-relevantan rizik = AUD-039** (shipped destruktivni test makroi, potvrđeno `modRelease.bas:5,18`): najoštriji je **`Test_StornoCentar_All` bez ijednog env guard-a (P1)** — 23 storno testa nad živim podacima na jedan Alt+F8, bez potvrde. Sekundarno: `RunStornoTestSuite`/`modTestPalete` imaju bar MsgBox; `modSEFTests` ima config-flag+kucanu potvrdu (najbolji).

4. **Jeftini konkretni correctness fixevi (S):** (a) `TBL_STORNO_ZURNAL` snapshot u `modTestStorno:58-68` i `Test_StornoSelectedBlocks_Auto:640` — jednolinijski, preživljava single-writer; (b) capture `Err.*` pre rollback-a; (c) residue-preflight za `modTestStorno`/`modTestStornoCentar` (preslikati iz `modTestPalete`).

5. **Sistemska klasa false-green** (potvrđuje kontekstni AUD): fail-soft helperi (`SvAppend`/`TcSeedRow` silent skip, `Smoke_FirstValue`→"", `CountSEFEventsForFaktura`→0) + assertion-brojači bez case-metrike + „...OK"/MsgBox bez Boolean-a. **Nijedan od 5 modula nije u `modE2EReleaseGate`**, pa je „E2E false-green" trenutno **latentan** (uslovan na buduću integraciju), ne aktivan.

6. **Preklapanja:** AUD-016 (duplirani/parni test moduli), AUD-039 (shipped bez guard-a + E2E false-green) pokrivaju većinu „P0" tema na nivou obrasca; per-fajl ostaju specifični jednolinijski fixevi (zurnal-snapshot) i `modTestStornoCentar` env-guard kao najprioritetnija pojedinačna stavka. FM-0101 je čist P2 repo-cleanup bez funkcionalnog rizika.

---

## v142 blok K6 — modComboBinding, frmMarza, frmExcelMini, modPaletniListUI, clsBlokUI, clsSEFValidationResult, clsWheelList, modMouseWheel (FM-0103/0106/0114/0116/0121/0123/0127/0140) [sidro origin/main v2.24.0]

Verifikacija kompletna. Sve tvrdnje su proverene protiv koda. Ključni kontekst koji oblikuje ocene:

- **Permission model** (`modAuth.bas`): opt-in preko `AuthEnabled()` (**OFF po defaultu**, red 17), a jedini enforcement je na navigaciji forme — `OpenContentForm` (`frmOtkupAPP.frm:1072-1076`) preko `OblastZaFormu`. Soft workflow-segregacija, ne hard security (xlsm nezaštićen).
- `EnsureUserFormChromeRemoved` postoji (`modWindow.bas:119`), `ExtractIDFromDisplay` postoji i **razlikuje se** od `ExtractIDFromDisplaySafe` (`modHelpers.bas:13`), `RequireColumnIndex` **baca grešku** (`modSchemaGuard.bas:13`), `frmPalete.frm` **postoji**, `clsSEFValidationResult` **nema nijednog VBA callera**.

---

### FM-0103 — modComboBinding.bas

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
|123.3|`.Value` vraća display, ne ID (`:29-30` BoundColumn=1/TextColumn=1)|P1|**Tačno** (mehanizam), ali namerno i **već dokumentovano** (`:10-13`); BoundColumn=2 bi slomio display contract|P3|Pojačati komentar/naming; NE dirati BoundColumn|S|
|123.4|Destruktivni `Clear` (`:26`) pre validacije kolona (`:40-41` RequireColumnIndex baca)|P1|**Tačno** — combo ostane prazan, greška se guta|P2|Validiraj izvor pre `Clear` ILI vrati broj/Boolean|S|
|123.5|Prazna tabela i schema-drift izgledaju isto (`:35` tihi Exit)|P1|**Tačno** (isti koren kao 123.4)|P3|Fill vrati status/broj dodatih|S|
|123.6|Filter aktivnog uzak (`:56-57` samo ==STATUS_NEAKTIVAN)|—|**Dizajnersko ograničenje** — default-active je bezbedniji za soft-delete (ne krij aktivne zbog typo)|P3|Dokumentovati; opc. `IsMasterRowActive`|S|
|123.7|Prazan display dozvoljen (`:60` samo `idValue<>""`)|—|**Tačno**|P3|Fallback display `[ID]` ili zahtevaj neprazan|S|
|123.8|Dupli ID nisu detektovani; SetComboByID uzima prvi (`:115-120`)|—|**Tačno**, ali ID je master PK → defanzivno|P3|Opc. duplicate-guard (dev/release)|S|
|123.9|Case-sensitive `=` (`:116`, `:145`)|—|**Delimično/Kontekst** — za numeričke ID-jeve (KupacID/KulturaID) irelevantno|P3|`vbTextCompare` radi konzistentnosti|S|
|123.10|Legacy fallback greši ID (`:163-167` InStrRev)|—|**Tačno** za edge-case; samo legacy fallback|P3|Ostaviti kao migraciju; novi combo 2-kol|S|
|123.11|Nekonzistentan parser: GetComboID→`...Safe` (`:84`) vs SelectComboByDisplayID→`ExtractIDFromDisplay` (`modHelpers.bas:13`, drugačija logika+fallback)|P1|**Tačno** — stvarna nekonzistentnost potvrđena|P2|Ujednačiti na `...Safe` ili dokumentovati razliku|S|
|123.12|Fill nema rezultat; Set/Select vraćaju False|—|**Tačno** (spoj sa 123.4)|P2|Fill vrati broj/Boolean|S|
|123.13|Nedostaje test matrica|—|**Nije proverivo statički**|P3|Opc. regresioni harness|M|

**Bilans:** Modul tehnički tačan i dobro ograničen; većina nalaza je Tačno-ali-nisko (P2/P3) jer je single-writer + `.Value/GetComboID` contract već dokumentovan. Najvredniji: **123.11** (ujednači parser) i **123.4/123.12** (Fill vrati status pre destruktivnog `Clear`). BoundColumn NE dirati.

---

### FM-0106 — frmMarza.frm  (LEGACY — korisnik potvrdio da se ne koristi)

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
|126.4|Naziv kao identitet (`:83`/`:117` Naziv→LookupValue→KupacID)|P1|**Tačno**|Prihvaćeno/P3 (legacy)|Ako se oživi: `FillComboDisplayID`/`GetComboID`|M|
|126.5|Pogrešan ID nije fail-fast (`:117-118`, nema `ID<>""` pre report)|P1|**Tačno**|P3 (legacy)|Validiraj `ID<>""` pre poziva|S|
|126.6|Datum locale-zavisan (`:108-109` CDate; nema `Od<=Do`)|P1|**Tačno**|P3 (legacy)|`TryParseDateValue` (FM-0102) + provera opsega|S|
|126.7|„Po OM" je procena pod generičkim „Marža"|P1|**Delimično / Nije proverivo** — metodologija u `modMarza` (nije čitan), headeri u `.frx`|P3 (legacy)|Label „Procena po OM" + metodologija|M|
|126.8|Različita semantika tri režima|P1|**Nije proverivo statički** (modMarza)|P3|Dokumentovati metodologiju|M|
|126.9|Ukupno bez inventory matching|P1|**Nije proverivo statički** (modMarza)|P3|Imenovati „period contribution"|M|
|126.10|Invalid datumi mogu u agregat|P1|**Nije proverivo statički** (modMarza helperi)|P3|Isključi+prijavi invalid|M|
|126.11|Schema guard regresija; forma daje generičku grešku (`:136`)|P1|**Delimično / Nije proverivo** (modMarza)|P3|`RequireColumnIndex` u modMarza|S/M|
|126.12|Nema pravi caption (`:3` „UserForm1"→`:46` „")|P2|**Tačno** — FindWindow (`:34`) sa praznim caption|P3|Postaviti caption pre FindWindow ILI central helper|S|
|126.13|Windows-only API (`:30-42`)|P2|**Dizajnersko** (Windows-only app)|Prihvaćeno|Opc. central helper|S|
|126.14|One-shot setup pri Hide (`:26`/`:51-52`; `:156` Me.Hide→stale)|P2|**Tačno**|P3 (legacy)|Razdvoji one-time i per-show|S|
|126.15|Lista može zadržati stare vrednosti (`:144`)|P2|**Tačno**|P3|Eksplicitno prazni ne-numeric ćelije|S|
|126.16|Zaglavlja/metodologija nisu u kodu (u `.frx`)|P2|**Nije proverivo statički**|P3|— (dizajner)|—|
|126.17|Hardcoded tekstovi van `Poruka()` (`:59-61`,`:128`)|P2|**Tačno** — odstupa od lokalizacione arhitekture|P3 (legacy)|Kroz `Poruka()` ako se oživi|M|

**Bilans:** Forma je legacy/nekorišćena → svi nalazi Prihvaćeno/P3. Tehnički tačni ili plauzibilni; najveći (procena vs ostvarena marža) leži u `modMarza` (nije čitan) i `.frx` headerima, pa nije statički potvrdiv iz forme. Bez oživljavanja nema poslovnog rizika; ako se oživljava: prvo stable-ID binding + validacija ID/perioda.

---

### FM-0114 — frmExcelMini.frm  (live: `frmOtkupAPP.frm:833` Show modeless)

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
|—|Moguće stanje bez UI: `:102` Visible=False PRE `:105` Show; EH (`:114-117`) ponavlja|P1|**Tačno** mehanizam; ali frmOtkupAPP je uvek-učitan shell (retko pada)|P2|Show→potvrdi→pa Visible=False; na grešku ostavi Visible=True|S|
|—|Greška pri `Show` se guta (`:104-106` On Error Resume Next)|P1|**Tačno** (isti koren)|P2|Proveri load/loguj source pre fallback|S|
|—|Ne vraća prethodni Excel state|P2|**Dizajnersko** (namenski kiosk return)|P3|Dokumentovati kiosk contract|S|
|—|Permission nije lokalni (OBL_OTVORI_EXCEL)|P2|**Delimično** — ulaz `btnOpenExcel_Click` GA proverava (`frmOtkupAPP.frm:821-825`) pre Show; lokalna provera u return-helperu promašena|P3|— (guard je na ulazu); opc. potvrditi sve Excel-visible puteve|S|
|—|Duplirana WinAPI chrome (`:30-42` vs `EnsureUserFormChromeRemoved`, modWindow:119)|P2|**Tačno** — anti-duplikacija (poklapa CLAUDE.md doktrinu)|P2|Koristi `EnsureUserFormChromeRemoved`|S|
|—|Pogrešna log oznaka (`:65`/`:82`/`:112` „frmCloseExcel.*" a VB_Name=frmExcelMini `:11`)|P2|**Tačno** — telemetrijski mismatch|P2|Ispravi labele na `frmExcelMini`|S|
|—|Designer caption „UserForm1" (`:3`)|P2|**Tačno**|P3|— ili central helper|S|

**Bilans:** Mali live helper (ulaz zaštićen OBL_OTVORI_EXCEL). Najvredniji jeftini: preurediti Show→hide redosled (P2), ispraviti log oznaku frmCloseExcel→frmExcelMini (P2), koristiti central chrome helper (P2, anti-dup). Ostalo P3.

---

### FM-0116 — modPaletniListUI.bas

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
|—|Alt+F8 destruktivni zaobilaze permission (nema `KorisnikImaPravo(OBL_PALETE)`; guard samo u `OpenContentForm`, frmOtkupAPP:1072)|P1|**Delimično** — mehanizam tačan KAD je AUTH ON; ali AUTH opt-in/OFF (modAuth:17), soft model, xlsm nezaštićen (Alt+F8/VBA ionako sve može)|P2 (defense-in-depth)|`If AuthEnabled() And Not KorisnikImaPravo(OBL_PALETE) Then...Exit` na svaki destruktivni entry (obrazac postoji u frmOtkupAPP)|M|
|—|`Val` prihvata delimično nevalidno (`:97-98`)|P1|**Delimično** — jednokratni prompt-ovi SU IsNumeric-guarded (`:19,44,129,154,184`) → „12abc" odbačen; nezaštićen `Val` samo u SavePrerada|P2|Strogi parser (modParse) u SavePrerada|S|
|—|Lista prerade prazna/duplikati (`:76-86`→`:97`)|P1|**Tačno** (UI-nivo): tokeni `pbr<=0` tiho preskočeni → prazan `ids` ide u TX; nema dedupe|P2|Zahtevaj ≥1 jedinstvenu paletu pre poziva|S|
|—|Samo tekuća godina (svi `Year(Date)`)|P1|**Tačno**; **Dizajnersko** — Alt+F8 stub, frmPalete postoji za pun tok|P2/P3|Birati PaletaID/PreradaID ili tražiti godinu|M|
|—|Success bez provere rezultata — Close (`:140-141`)|P1|**Tačno** — `ClosePaletaManual_TX` vraća String (modPaletniList:595) ali se ignoriše; nekonzistentno (Storno* proveravaju `:167,197`)|P2|Uhvati povratni String, potvrdi post-condition|S|
|—|Export rezultat se ne proverava (`:31,56,99`)|P1|**Tačno** (nema post-provere), ali export ima svoj EH|P3|Opc. potvrda generisanja|S|
|—|`PaletaAdjustPrompt` vraća tekst (`:232-248`)|P1|**Delimično** — koristi se za MsgBox prikaz (`:260`); nema automatizovanih callera → struct je over-engineering|P3|— (ostaviti); struct tek uz automatizovan caller|M|
|—|Težina lookup vraća 0 za missing/error (`:288-303`)|P2|**Tačno** (0 dvosmisleno)|P3|Typed/Boolean lookup|S/M|
|—|Legacy Alt+F8 sloj posle frmPalete (komentar `:8-9`)|P2|**Tačno** — frmPalete.frm postoji|P2 (anti-dup)|Zadrži legit admin/recovery; ostale private/ukloni|M|

**Bilans:** Tanak adapter, business delegiran. Nalazi stoje ali ublaženi: permission-bypass je P2 defense-in-depth (ne P1) zbog AUTH opt-in/OFF + soft modela, a `Val` nalaz je preširok (prompt-ovi IsNumeric-guarded). Najvredniji jeftini: ClosePaleta uhvati rezultat, ≥1 jedinstvena paleta, guard-obrazac na destruktivne Alt+F8 kad je AUTH ON. frmPalete postoji → očistiti legacy stubove.

---

### FM-0121 — clsBlokUI.cls

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
|—|Destruktivni router bez permission: `OtkupBlok_OnButton` „STORNO" (modOtkupBlok:134,138)|P1|**Delimično** — router zavisi od ŽIVE frmOtkup+selektovanog bloka (mForm); naked Alt+F8 no-op/greška. Živ blok znači frmOtkup već otvoren (prošao OBL_OTKUP) → inkrementalni bypass ~nula; + AUTH opt-in|P3|Opc. domain-guard u `StornoSelectedBlok`/Storno_TX|S/M|
|—|String-based routeri fail-silent (nema Case Else)|P2|**Tačno** (klasa prosleđuje `action`; Select Case u modOtkupBlok)|P3|`Case Else`→LogErr; centralne konstante|S/M|
|—|`action` nije tipiziran po kontroli|P2|**Dizajnersko** — mala runtime klasa; FM sam kaže refactoring nije obavezan|P3|Validirati kombinaciju pri wiring-u|S|
|—|Javna polja → polu-inicijalizovan wrapper (nema Init)|P2|**Tačno**|P3|Init: tačno jedna kontrola + neprazan action|S|
|—|Slab error context|P2|**Tačno**|P3|Logovati action+control name|S|
|—|`txt_Change` učestali side-effect (`:33`)|P2|**Kontekst-Pozitivno** — CENA grana samo osvežava state; persistence je u AfterUpdate|P3|Zadržati Change bez skupih upisa|—|
|—|Globalno module-state vezivanje (jedan mForm)|P2|**Dizajnersko ograničenje** — single frmOtkup instanca|P3|Dokumentovati single-instance contract|S|

**Bilans:** Dobar mali WithEvents adapter bez business logike. STORNO-bypass ublažen (router zavisi od žive forme koja je već prošla OBL_OTKUP; AUTH opt-in) → P3, ne P1. Ostalo P3 dijagnostika/robustnost. Rebuild-stale-sink nije problem jer `mWrappers` drži instance žive. Minimalno vredno: `Case Else` fail-fast + log action/control.

---

### FM-0123 — clsSEFValidationResult.cls

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
|—|Prazna metadata-only klasa, bez callera|P2|**Tačno** — 9 linija, bez članova; grep: samo `.cls` + docs/changelog/backlog, nijedan VBA caller|P2 cleanup|Opcija A: obrisati `.cls`|S|
|—|Lažna arhitektonska slika (naziv sugeriše typed-result)|—|**Tačno**|P2|Ukloniti changelog/backlog reference koje tvrde typed-result|S|
|—|Import/build šum (prazna klasa se uvozi)|—|**Tačno**|P3|Brisanje uklanja VBComponent|S|
|—|Opcija B — implementirati typed result|—|**Prihvaćeno** — samo uz migracioni plan; FM se slaže da se NE radi unapred|—|— (ostaviti za backlog)|—|

**Bilans:** Potpuno tačno — mrtva prazna klasa bez callera; jedini „rizik" je misleading arhitektura + import šum. Brisanje (Opcija A) je ispravan minimal-delta; NE praviti placeholder „za kasnije". Uz brisanje očistiti changelog/backlog tvrdnje o typed-result. P2, S.

---

### FM-0127 — clsWheelList.cls

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
|—|Jedini handler guta greške (`:30-32` On Error Resume Next)|P2|**Kontekst-Pozitivno / Dizajnersko** — fail-soft na hover je namerno ispravan (ne sme oboriti formu)|P3|Dev-režim: log prvog failure-a po wrapperu|S|
|—|Javna kontrolna referenca promenljiva (`lst` Public)|P2|**Tačno**; modMouseWheel to ne zloupotrebljava|P3|Opc. `Attach(lb)`+private ref (nije nužno)|S|
|—|Nema eksplicitni Detach|P2|**Dizajnersko** — globalni `MouseWheel_Detach` dovoljan za all-or-nothing|P3|Detach tek ako se pojedinačne liste uklanjaju|S|
|—|Komentar „bez stanja" nepr.  (drži COM ref)|P2|**Tačno** — samo dokumentaciona preciznost, nije problem|P3|Preformulisati komentar|S|
|—|Owner/form identitet nije sačuvan|P2|**Dizajnersko** — dobro za minimalizam|P3|Opc. dev-metadata FormName/ControlName|S|

**Bilans:** Kvalitetan minimalan adapter; FM sam potvrđuje da su glavni rizici u modMouseWheel, ne ovde. Svi nalazi P3 dijagnostika/preciznost; fail-soft na hover je namerno i ispravno. Ne širiti klasu. Jedino opciono: dev-log prvog SetHot failure-a.

---

### FM-0140 — modMouseWheel.bas  (off-by-default, opciona kozmetička UI funkcija)

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
|—|Windows timeout → lažno aktivan `mHook` (`:169` `If mHook<>0 Then Exit`)|P0/P1|**Delimično** — komentar-mismatch stvaran (`:16-17`), ALI modul ima sopstveni keep-alive koji sam unhook-uje (`:282-287`), a Detach na Deactivate/QueryClose resetuje mHook → self-recovery na re-fokus („trajno mrtav" preterano)|P3 (kozmetička/off-default)|Uskladiti komentar (S) ILI `mLastCallbackAt` heartbeat (M)|S/M|
|—|`UnhookWindowsHookEx` rezultat ignorisan (`:114-115`, `:284-285`)|P0/P1|**Tačno** — kod ignoriše povratnu vrednost, briše handle bezuslovno. Isti-thread unhook praktično ne pada, ali dangling AddressOf→crash je dovoljno ozbiljan (jedini realan crash-vektor)|**P2** (najlegitimniji nalaz grupe)|`If UnhookWindowsHookEx(h)<>0 Then mHook=0 Else zadrži+loguj`|S|
|—|Globalni detach nije owner-aware (`:138-140` vs `:111-119`)|P1|**Dizajnersko ograničenje** — single-form modalni UX (FM priznaje); self-heal na sledeći Activate|P3|Dokumentovati single-form contract (S); owner-registry je over-engineering (L)|S|
|—|`mHot` vezan za uklonjenu listu|P1|**Delimično** — COM ref drži živo (nije use-after-free); ScrollHot guardovan (`:295-298`)|P3|Opc. unregister po kontroli|M|
|—|On Error Resume Next skriva install-fail (`:167,178`)|P1|**Tačno**; ali fail-soft ispravan za opcionu funkciju (FM priznaje)|P3|Log jednom po sesiji na install-fail|S|
|—|VBE guard može trajno onemogućiti (`:184-190`→True)|P1|**Kontekst/Dizajnersko** — fail-closed nameran+bezbedan; app ionako traži „Trust access VBA object model" za self-update → scenario nizak|P3|Dijagnostički razlog (VBE-open vs unavailable)|S/M|
|—|Komentar „gasi se čim VBE otvoren" jači od impl (`:176` samo u EnsureHook)|P1|**Delimično** — normalno otvaranje VBE→forma gubi fokus→Deactivate→Detach; ostaje egzotičan multi-monitor edge|P3|Uskladiti komentar|S|
|—|Auto-unhook broji sve mouse događaje (`:282`)|P1|**Kontekst-Pozitivno** — dizajn ok (SetHot vraća mAlive na pun `:160`); samo komentar „24 pokreta" nepr.|P3|Preformulisati komentar|S|
|—|Delta kao znak (`:273-274` samo `md>0`)|P2|**Dizajnersko** — za standardni miš prihvatljivo|P3|Opc. signed high-word akumulacija do 120|M|
|—|Clamp `n-1` (`:307`)|P2|**Tačno**; nije kritično|P3|Opc. procena vidljivih redova|M|
|—|Javni Alt+F8 On/Off/Reset (`:132,143,149`)|P2|**Prihvaćeno** — namerni dijagnostički entry (`:45,129`)|Prihvaćeno|Opc. `Option Private Module` (gubi dijagnostiku)|S|
|—|Nema health/debug API|P2|**Tačno** (samo Reset)|P3|Opc. status funkcija|M|
|—|32/64-bit + Mac|—|**Kontekst-Pozitivno** (PtrSafe/LongPtr ok, MD_OFFSET=8 obrazložen); Mac: `#If VBA7` aktivan i na Mac-u → Win Declare pao, ali app Windows-only|Prihvaćeno|Formalni Windows-only build guard ili `#If Mac`|S|

**Bilans:** Iznadprosečno pažljiv off-by-default hook; većina „P0/P1" oznaka je **napumpana** za kozmetičku opcionu funkciju bez data/crash rizika (scrollbar radi i bez nje). Jedini stvarno vredan jeftini hardening: **proveriti `UnhookWindowsHookEx` rezultat i ne gubiti handle (P2, S)** — pokriva jedini realni crash-vektor (dangling AddressOf). Stale-hook heartbeat i owner-registry su korisni ali za single-form/off-default scope P3/over-engineering. Komentar-vs-impl neusklađenosti (timeout re-hook, „gasi se čim VBE", „24 pokreta") su stvarne ali kozmetičke — uskladiti tekst.

---

## Zbirni zaključak audita

**Najviše precenjeno (FM P0/P1 → realno P2/P3):** modMouseWheel timeout/lifecycle (kozmetička off-by-default funkcija), permission-bypass nalazi (modPaletniListUI/clsBlokUI — AUTH je opt-in/OFF + soft model + nezaštićen xlsm), sve frmMarza P1 (forma legacy/nekorišćena).

**Najvredniji stvarni, jeftini nalazi (minimal-delta, potvrđeni protiv koda):**
1. **modComboBinding 123.11** — dva različita ID-parsera (`ExtractIDFromDisplay` vs `...Safe`); ujednačiti. S.
2. **modMouseWheel** — proveriti `UnhookWindowsHookEx` rezultat, ne gubiti handle (jedini realan crash-vektor). S.
3. **frmExcelMini** — log oznaka `frmCloseExcel`→`frmExcelMini` + central `EnsureUserFormChromeRemoved` (anti-dup po CLAUDE.md). S.
4. **modPaletniListUI** — `ClosePaleta_Prompt` uhvati povratni String (nekonzistentno sa Storno*). S.
5. **clsSEFValidationResult** — obrisati mrtvu klasu + changelog/backlog reference. S.

**Netačnih/preširоких tvrdnji:** modPaletniListUI „Val" nalaz je preširok (jednokratni prompt-ovi su IsNumeric-guarded); modMouseWheel „trajno mrtav hook" i „gasi se čim VBE otvoren" su preterani (self-recovery preko Deactivate/Detach). Nijedan nalaz nije potpuno **Netačan** — svi imaju tačno mehaničko jezgro, ali kontekst (single-writer desktop, opt-in AUTH, off-by-default hook, legacy forma) sistematski spušta hitnost.

---

## v142 blok K7 — banka parseri: modBankaProCredit, modBankaAlta, modBankaHalk, clipboard(mrtav), modBankaImportParserPdfToText (FM-0128..0132) [sidro origin/main v2.24.0]

Verifikacija kompletna. Imam sve module (5 parsera + orkestrator `modBankaImport.bas`) i potvrđen ključni kontekst koji FM skoro potpuno ignoriše.

---

## Auditski nalaz — FM-0128…FM-0132 (banka PDF parseri)

**Zajednički kontekst koji presuđuje većinu nalaza (FM ga ne pominje):** `modBankaImport.ParseBankaIzvodForImport` (`modBankaImport.bas:485-546`) primenjuje **4-nivo fail-closed integritet PRE staging-a**, za SVE banke posle dispatch-a: Lvl1 `Pocetno+Uplate-Isplate==Novo`, Lvl2 `sumUplata==Potrazuje`, Lvl3 `sumIsplata==Duguje`, Lvl4 `countUplata==nOdobrenje` i `countIsplata==nZaduzenje`; svaki promašaj `Err.Raise` → **ceo izvod se odbija** (nema tihe korupcije). Uz to `saldo.parsed` je obavezan (`:473`). **Posledica:** promašen/izmišljen iznos, procureli neizvršeni nalog, oba-smera red, layout-drift — sve to razbija sume/brojeve i puca fail-closed. **Jedina rupa koju gate DOKAZIVO ne vidi:** red sa `zad=0 And odo=0` (ne menja ni sume ni count-ove) — ali je novčano nulti. Ovo je osa po kojoj obaram FM ocene sa „tiha korupcija" na „prekid uvoza / novčano-nulti phantom". Verifikovano i: clipboard parser ima **0 produkcionih callera** (grep — samo sopstveni `TestParser`); PdfText je **živ** (Komercijalna = `Case Else`, `:452-457`).

Pravilo ocene: **Tačno** = mehanizam i rezidualni rizik oba stoje (gate ne pokriva); **Delimično** = mehanizam tačan ali FM precenjuje rizik jer gate hvata na nivou izvoda.

---

### FM-0128 — `modBankaProCredit.bas`

| # | Nalaz (FM) | FM tež. | Opravdanost | Hitnost | Predlog (min-delta) | Napor |
|---|---|---|---|---|---|---|
| — | Pozitivni nalazi (izolovan, blok-bound do sledećeg RB, phantom filter, full-match račun/iznos) | — | **Kontekst-Pozitivno** (`:203-221,239-253`) | — | — | — |
| 1 | Fin. polja na fiksnim offsetima pivot+1/+3/+5 | P1 | **Delimično** — mehanizam tačan (`:300-308`), ali „opasno/red validan" precenjeno: gate hvata i promašen (Lvl2/3/4) i izmišljen iznos | P2 | Skenirati `IsAmount` u prozoru [pivot+1..sifra] umesto fiksnih indeksa | M |
| 2 | `IsRealTxnPC` prima account-only 0/0 | P1 | **Tačno** (`:241-253`) — gate 0/0-slep (Lvl4 count-blind na nule) | P2 | Pooštriti: zahtevati `zad>0 Xor odo>0` (deljeno s Alta #1) | S |
| 3 | Moguća oba iznosa pozitivna | P1 | **Delimično** — nema lokalne invarijante, ali Lvl2/3/4 hvataju | P3 | Opc. `Not(zad>0 And odo>0)` radi preciznije poruke | S |
| 4 | Date validator shape-only | P1 | **Tačno** (`:391-394`) — gate ne validira datum; `31.02` ulazi kao string | P2 | Kalendarska validacija u zajedničkom `modBankaParseUtils` (RF-16, ×5) | S |
| 5 | Svaki standalone datum = pivot (bez sekcijskog bound-a) | P1 | **Tačno** (`:178-183`) — ProCredit JEDINI nije sekcijski bound (Alta/Halk jesu) | P2 | Bound pivote na „STANJE I PROMENE"/PROMENE kao Alta | M |
| 6 | `FindRbAbovePC` pada do `LBound` | P1 | **Tačno** (`:345-354`) — posledica: prenapuhan partner blok | P3 | Ako nema RB u prozoru (~8 lin.) → kandidat nevalidan | S |
| 7 | Račun bez checksum/normalizacije | P1 | **Tačno** (`:438-444`) — ali partner-konto je metapodatak, van gate-a | P3 | Centralna normalizacija/opc. mod-97 (deljeno ×5) | M |
| 8 | Poziv na broj parser preuzak | P1 | **Tačno** (`:467-477`) — poziv je jak ključ za auto-map | P2 | Proširiti obrazac (model `(NN)`, crtice, duži) — deljeno s Alta/Halk | M |
| 9 | Referenca = prvi ≥6 cifara (heuristika) | P1 | **Tačno** (`:311-318`) — ref ulazi u dedup (`IsDuplicateBankaImport` jak ključ) | P2 | Kontekstualnija potvrda; sačuvati raw svrhu | M |
| 10 | Saldo rigidno 6 tokena, samo Boolean | P2 | **Tačno** (`:356-386`) — fail-closed; orch. daje jasnu 1003 | P3 | Reason-code opcionalno | S |
| 11 | Error contract `Empty`/`parsed=False` | P2 | **Delimično** — orkestrator diže bogate typed greške 1000-1008 (samo per-red razlog fali) | P3 | Per-kandidat reason ako se uvodi typed rezultat | M |
| 12 | Regex u hot path-u | P2 | **Tačno** (`:396-402` i sl.) — perf, ne rizik | P3 | Keširati module-level `RegExp` (deljeno ×5) | S |

**Bilans FM-0128:** Strukturni opis tačan i koristan; FM ozbiljno precenjuje rizik jer ignoriše 4-nivo integritet — nema tihe korupcije, drift = prekid uvoza. Stvarni rezidual koji gate ne vidi: 0/0-sa-računom (#2, novčano nulti). ProCredit je zapravo slabiji od Alta/Halk po jednoj tački (#5 — bez sekcijskog bound-a), što FM ispravno hvata. Realni prioriteti: **P2** #2/#4/#5/#8/#9; ostalo P3. #4/#7/#12 su deljeni → RF-16.

---

### FM-0129 — `modBankaAlta.bas`

| # | Nalaz (FM) | FM tež. | Opravdanost | Hitnost | Predlog (min-delta) | Napor |
|---|---|---|---|---|---|---|
| — | Pozitivni (sekcijski bound PROMENE..Ukupno, fee anchor semantički, zasebni uplata/isplata regex) | — | **Kontekst-Pozitivno** (`:198-215,407-448`) | — | — | — |
| 1 | Account-only 0/0 prolazi | P1 | **Tačno** (`:273-281`) — identično PC #2; gate slep | P2 | Zahtevati `zad>0 Xor odo>0` (isti fix kao PC) | S |
| 2 | Moguća oba smera pozitivna | P1 | **Delimično** — Lvl2/3/4 hvataju | P3 | Opc. lokalni guard | S |
| 3 | Svaki kratki broj = R.B. | P1 | **Delimično** (`:485-490`) — bound na PROMENE (bolje od PC); unutar sekcije rizik blok-splita, ali kvar iznosa → gate | P2 | Potvrditi kontekst (sledeće lin. partner/račun) / rastuću sekvencu | M |
| 4 | Datum shape-only | P1 | **Tačno** (`:492-495`) — deljeno | P2 | RF-16 kalendarska validacija | S |
| 5 | Txn bez datuma knjiženja prolazi | P1 | **Tačno** (`:336-344,273-281`) — `IsRealTxnAlta` ne gleda datum | P2 | Dodati validan datum u `IsRealTxnAlta` | S |
| 6 | Šifra-linija nije obavezna | P1 | **Delimično** (`:368-377`) — ako je uplata, propušten odobrenje → Lvl2/4 puca | P2 | Missing šifra-linija = strukturirani warning | S |
| 7 | Fee anchor izostanak ne obara red | P1 | **Delimično** (`:346-363`) — menja skenove; promašaj iznosa → gate | P3 | Razlikovati layout-varijantu od kvara | M |
| 8 | Uzima prvi račun u bloku | P1 | **Tačno** (`:297-303`) — metapodatak, nizak uticaj | P3 | Kontekstualna pozicija + normalizacija (deljeno) | M |
| 9 | Poziv pogrešno izvučen iz svrhe | P1 | **Delimično** (`:570-591`) — auto-map je predlog koji operater potvrđuje | P2 | Sačuvati raw svrhu/blok | M |
| 10 | Referenca = poslednji 15-cifreni token | P1 | **Tačno** (`:551-560`) — ulazi u dedup | P2 | Field-level potvrda; raw tail | M |
| 11 | Saldo tačno 6 tokena | P2 | **Tačno** (`:453-483`) — fail-closed | P3 | Reason-code | S |
| 12 | `Empty` ne razlikuje scenarije | P2 | **Delimično** — orch. typed greške | P3 | — | M |
| 13 | PROMENE/Ukupno fallback fail-OPEN | P2 | **Tačno** (`:200-215`) — `txStart=LBound`/`txEnd=UBound` ako anchor fali | P2 | Missing anchor → parse failure (fail-closed), ne ceo dokument | S |

**Bilans FM-0129:** Alta strukturno najotporniji (sekcijski bound + fee anchor) — FM to priznaje. Isti gate mitigira oba-smera/promašaj/procurivanje. Realni rezidual: 0/0-sa-računom (#1), prazan datum knjiženja (#5), fail-open sekcijski anchor (#13, jedina stvarna P2 lokalna zaštita). RF-16 deli #4/#8/#11.

---

### FM-0130 — `modBankaHalk.bas`

| # | Nalaz (FM) | FM tež. | Opravdanost | Hitnost | Predlog (min-delta) | Napor |
|---|---|---|---|---|---|---|
| — | Pozitivni (NEIZVRŠENI isključenje, `N.` specifičan, `Poreklo naloga` start, context-bound saldo) | — | **Kontekst-Pozitivno** (`:191-224,161-178`) — saldo bolji od PC/Alta (3 anchora, ne globalni) | — | — | — |
| 1 | Svaki R.B. bezuslovno ulazi (nema `IsRealTxnHalk`) | **P0/P1** | **Tačno** (`:231-249`) — Halk JEDINI bez `IsRealTxn`; ali phantom je novčano nulti (gate 0/0-slep) → **P0 precenjeno** | **P1** | Dodati `IsRealTxnHalk` pre copy (poravnati s PC/Alta) | S |
| 2 | Sekcijski anchori fail-OPEN | P1 | **Tačno** (`:193-213`) — `executedEnd=UBound`/`executedStart=LBound` | P2 | Obavezni `Poreklo naloga` + `Ukupno` | S |
| 3 | Izvršeni/neizvršeni se mešaju ako `Ukupno` izostane | P1 | **Delimično** (`:195-201`) — procureli nalozi razbijaju Lvl2/3/4 → import blokiran (fail-closed) | P2 | Tvrdi stop i na `NEIZVRSENI NALOZI` | S |
| 4 | Zaduženje fiksni `afterDates+1` | P1 | **Tačno** (`:314-317`) — krući od Alte; promašaj → gate | P2 | Prozor do anchora kao Alta | M |
| 5 | Txn bez datuma prolazi | P1 | **Tačno** (`:305-312`) — + nema `IsRealTxn`; obično 0/0 red | P2 | Datum obavezan u `IsRealTxnHalk` | S |
| 6 | Datum shape-only | P1 | **Tačno** (`:370-373`) — deljeno | P2 | RF-16 | S |
| 7 | Nema invariant smera | P1 | **Delimično** — gate | P3 | — | S |
| 8 | `ParseOdoSifHalk` prima bilo koju `<iznos> <3cifre>` | P1 | **Delimično** (`:383-396`) — prvi match posle datuma; pogrešan iznos → gate | P2 | Vezati za fee/poziciju | M |
| 9 | Račun nije obavezan, bez checksum | P1 | **Tačno** (`:266-287,419-425`) — metapodatak | P3 | Deljena normalizacija | M |
| 10 | Referenca prekida svrhu na 1. 13-cifrenom | P1 | **Tačno** (`:334-349`) — gubi trailing svrhu + moguća pogrešna ref (dedup) | P2 | Sačuvati raw tail; kontekstualna ref | M |
| 11 | Saldo fail-closed bez razloga | P2 | **Tačno** (`:438-466`) — ali context-bound (bolji) | P3 | Reason-code | S |
| 12 | Matrica bez confidence/source-range | P2 | **Delimično** — orch. typed greške | P3 | — | M |

**Bilans FM-0130:** Halk ima najslabiji završni contract — JEDINI bez `IsRealTxn`, fiksni `afterDates+1`, fail-open sekcija; FM to tačno hvata. Ali **P0 je precenjeno**: gate + obavezni saldo + single-writer = fail-closed za sve novčano; rezidual = novčano-nulti phantom red. Realni: **P1** dodati `IsRealTxnHalk` (jeftino, poravnava sa siblinzima) + tvrdi NEIZVRŠENI/sekcijski stop; P2 prozor za zaduženje (#4) i raw tail (#10).

---

### FM-0131 — `modBankaImportParserClipboard.bas` (MRTAV — RF-01, 0 callera potvrđeno)

| # | Nalaz (FM) | FM tež. | Opravdanost | Hitnost | Napomena |
|---|---|---|---|---|---|
| 1 | Svaki blok bezuslovno ulazi | P1 | **Tačno** (`:88-105`) | Prihvaćeno (RF-01) | Bespredmetno; ISTI obrazac živi u PdfText (FM-0132 #8) |
| 2 | `IsTxnStart` preširok `^\d{1,3}(\s+.*)?$` | P1 | **Tačno** (`:108-124`) | Prihvaćeno | Bespredmetno; PdfText `IsPdfTextTxnStart` je uža varijanta istog |
| 3 | Bez start-anchora + **duplirani OR** uslov | P1 | **Tačno** (`:62-63` — identičan `InStr…"Ukupno za ra"&ChrW(269)&"un"` ×2) | Prihvaćeno | **Zajednički uvid:** isti copy-paste dupli-OR ŽIVI u PdfText (`:364-365,484-485,503-504,676-677`) |
| 4 | Oba iznosa = 2 fizičke linije | P1 | **Tačno** (`:200-220`) | Prihvaćeno | = PdfText fiksne pozicije |
| 5 | Oba smera 0 / oba poz | P1 | **Tačno** | Prihvaćeno | = shared |
| 6 | Datum shape-only | P1 | **Tačno** (`:255-260`) | Prihvaćeno | = shared RF-16 |
| 7 | Partner heuristika | P1 | **Tačno** (`:163-192`) | Prihvaćeno | Bespredmetno |
| 8 | Ref/poziv menjaju svrhu, bez raw | P1 | **Tačno** (`:320-345`) | Prihvaćeno | = shared |
| 9 | `IsReference` ≥12 numeric | P1 | **Tačno** (`:425-432`) | Prihvaćeno | = PdfText `IsReferencePdf` identičan (`:711-714`) |
| 10 | `izvodBroj` parsira se, ne vraća | P2 | **Tačno** (`:28-37,93`) | Prihvaćeno | Zastareo contract |
| 11 | Nema error contract | P2 | **Delimično** | Prihvaćeno | Orch. bi ionako pokrivao |
| 12 | `TestParser` javni | P2 | **Tačno** (`:524`) | Prihvaćeno | Ceo modul se briše |

**Bilans FM-0131:** Svi nalazi tačni ali **bespredmetni posle RF-01** (mrtav, 0 callera — grep potvrdio). Jedina prenosiva vrednost: potvrđuje da „unconditional copy + preširok `IsTxnStart` + shape-only date + `IsReference≥12` + **dupli-OR `Ukupno` uslov**" NISU clipboard-specifični — identično žive u `modBankaImportParserPdfToText` (aktivan, Komercijalna). **Hardening usmeriti na PdfText, ne na clipboard.** Preporuka brisanja se slaže s RF-01; pre brisanja sačuvati fixture uzorke.

---

### FM-0132 — `modBankaImportParserPdfToText.bas` (AKTIVAN — Komercijalna `Case Else`)

Dve odgovornosti: **extraction transport** (cross-bank; `ExtractTextFromPdf`/`PickPdf`/`WriteAllTextUtf8`/`BankIzvodSaldo` — koriste ih SVI parseri) + **Komercijalna parser**. Extraction nalazi NISU integritet-zaštićeni (pre-parse shell/temp).

| # | Nalaz (FM) | FM tež. | Opravdanost | Hitnost | Predlog (min-delta) | Napor |
|---|---|---|---|---|---|---|
| — | Pozitivni extraction (path-check, wait+exitcode, non-zero≠prazan, stale-temp brisan, output-must-exist, UTF-8 stream, re-raise) | — | **Kontekst-Pozitivno** (`:42-100`) | — | — | — |
| 1 | Plain exe ime preko PATH-a | P1 | **Dizajnersko ograničenje** (`:131-140`) — kod to NAMERNO dozvoljava (PATH deploy); DEFAULT je apsolutna workbook-rel. putanja; single-writer | P3 | Setup health-check da preferira apsolutnu + test-command | S |
| 2 | Postojanje ≠ integritet (hash/sig/verzija) | P1 | **Dizajnersko ograničenje** (`:133-140`) — van threat-modela (operater sam instalira Poppler) | Prihvaćeno | Opc. verzija/test-command u `CheckPdfToTextExists` | M |
| 3 | Exit=0 + fajl može biti prazan/no-text-layer | P1 | **Delimično** (`:76`) — ishod fail-closed: prazan→parser Empty→orch. 1000; `Diag_DumpFullPdfText` već upozorava (`:1212`) | P2 | `If Len(Trim$(txt))=0 → Err.Raise` jasna „no text layer/OCR" | S |
| 4 | Temp sa bank-podacima može ostati (`On Error Resume Next`) | P1 | **Tačno** (`:211-219`) — podatkovna higijena; default lokalni TEMP | P2 | Best-effort log-warning na cleanup fail (ne obarati uvoz) | S |
| 5 | `APP_TEMP_PATH` može biti shared/network | P1 | **Delimično** (`:226-246`) — default `Environ$("TEMP")` lokalni; samo misconfig | P3 | Health-check/dok. da bude user-local | S |
| 6 | Temp folder `MkDir` samo 1 nivo | P1 | **Tačno** (`:241-243`) — operativno, ne security (FM se slaže) | P3 | Recursive create / setup pre-kreira | S |
| 7 | Cmd/putanje u error opisu | P1 | **Tačno** (`:64-69`+`LogErr`) — nisko lokalno; rizik samo remote monitoring | P3 | Redigovati u remote monitoringu | S |
| 8 | Bezuslovno kopira svaki blok (nema `IsRealTxn`) | P1 | **Tačno** (`:304-322`) — = Halk klasa; 0/0 phantom slip | P2 | `IsRealTxn` pre copy (deljeno) | S |
| 9 | `IsPdfTextTxnStart` svaki numerički ≤3 | P1 | **Delimično** (`:389-395`) — stop-anchori + gate | P2 | Kontekst/sekcija | M |
| 10 | Fiksne pozicije `secondDate+1/+2/+3` | P1 | **Tačno** (`:457-459`) — = PC klasa; drift→gate | P2 | Skenirati prozor umesto fiksnih | M |
| 11 | Nedostajući datumi → pogrešni indeksi (0-based) | P1 | **Tačno** (`:416-459`) — bez datuma `idxZad=1/idxNak=2/idxOdoSif=3` čitaju prve linije; obično 0/0 phantom | P2 | Missing date → blok nevalidan | S |
| 12 | Datum izvoda bez validacije | P1 | **Tačno** (`:325-342`) — jedan po izvodu, string; nizak uticaj | P3 | Kalendarska provera (RF-16) | S |
| 13 | Unique temp `Randomize/Rnd` + otkriva PDF base-name | P2 | **Tačno** (`:158-166`) — kozmetički/minor | P3 | GUID bez base-name | S |
| 14 | Modul prevelik, meša transport+parser (>1200 lin.) | P2 | **Tačno** (1263 lin.) — ali transport je VEĆ deljena infra (ripple pri splitu); + mrtvi helperi (`Find*`, `IsPdfAmountSifraLine`, `ParsePdfAmountSifraLine`, `NormalizePdfTextTxnStart` — neupotrebljeni) | P3 | Split transport uz RF-16 kad se modul sledeći put dira; obrisati mrtve helpere | L |

**Bilans FM-0132:** Extraction sloj je stvarno solidan (FM tačno hvali) — glavni „security" nalazi (#1/#2) su na single-writer desktopu **dizajnerska ograničenja**, ne aktivne pretnje; #3/#4 su realni ali fail-closed/higijenski (P2, jeftino). Parser sloj deli tačno iste rupe kao clipboard/PC/Halk (#8/#10/#11) — gate ih pokriva za novac, rezidual = 0/0 phantom. **Bonus (van FM):** živi copy-paste dupli-OR uslovi (`:364-365,484-485,503-504,676-677`) i grupa mrtvih helpera — čiste uz #14.

---

## Zajednička konsolidacija (poveži — „isti nalaz u više banaka")

- **Shape-only datum** (`IsDateLine*` × PC/Alta/Halk/Pdf/Clip = 5) → jedan kalendarski validator u **`modBankaParseUtils` (RF-16)**.
- **`RegExp` u hot-path-u** (`IsAmount*`/`IsAccountLine*`/`ParseSifra*` ×5) → module-level keš → RF-16.
- **Račun bez checksum/normalizacije** (`^\d{3}-\d{5,20}-\d{2}$` ×5) → deljena normalizacija.
- **Saldo „tačno 6 tokena, samo Boolean"** (PC/Alta/Halk = 3) — **NAPOMENA:** Komercijalna/`TryParseSaldoDataLine` (`:996-1089`) je robusniji (4-6 tokena, collapsed-zero logika), pa se P2 saldo-nalaz NE odnosi na PdfText.
- **`IsRealTxn` nedostaje/preslab:** Halk (nema uopšte), PdfText/Clipboard (nema), PC/Alta (imaju ali 0/0-sa-računom prolazi) → jedan deljeni „pravi red" validator (`zad>0 Xor odo>0` + validan datum) je najveći ROI, poravnava sva 4 živa parsera.
- **Dupli-OR `Ukupno` copy-paste bug:** živ u PdfText (4 mesta) + clipboard — kozmetički, čisti se uz RF-16.

**Ključna korekcija FM-a:** FM kroz svih 5 unosa piše kao da parsed redovi teku pravo u staging bez validacije. Netačno — 4-nivo fail-closed integritet (`modBankaImport.bas:485-546`) obara svaki novčano-relevantan kvar na nivou izvoda. Zato nijedan nalaz nije stvarni P0, a većina P1 je P2/P3. Jedini dokazani rezidual koji gate ne vidi je **0/0 red** (novčano nulti) — otud je najracionalniji jedini P1 posao dodati/pooštriti deljeni `IsRealTxn`.

---

## v142 blok K8 — modPoruke, modTheme, modStornoWarm (FM-0133/0134/0137) [sidro origin/main v2.24.0]

Verifikacija kompletna. Sav relevantni kod i unakrsne reference su pregledani. Sledi audit.

---

## Ključne unakrsne provere (osnov za ocene)

- **modPoruke**: `tblPoruke` je app-vlasništvo, kreiran sa **imenovanim** kolonama (`EnsureDataTable ... Array(COL_POR_KLJUC, COL_POR_TEKST)`, `modSetup.bas:1108`) i **very-hidden** (`ws.Visible = xlSheetVeryHidden`, `modSetup.bas:1116`) → operater ga ne edituje ručno. `LogErr → LogError` piše u dnevni fajl i **ne poziva `Poruka`** (`modLogError.bas:97-105,32`) → logovanje iz `Poruka.EH` je bez rekurzije (FM-ova bojazan neosnovana).
- **modStornoWarm**: jedini potrošač `GetWarmStornoDocs` je `frmDokumenta.PopulateFindResults` (`frm:5111`) koji kolekciju **samo čita** (gradi zaseban `hits`, `5121-5131`), resetuje `m_fnAllDocs=Nothing` pri otvaranju (`4998`), a elementi su **Variant nizovi** (value-copy). Invalidaciju zove **samo** `clsTransaction.CommitTx` (`cls:79`). `ThisWorkbook.Workbook_BeforeClose` → `modMain.ShutdownApp` (`doccls:78`) → `StopStornoWarm` (`modMain.bas:239`). Nekvalifikovani OnTime string je **konvencija celog projekta** (`modStanicaLock:238/248`, `modJournaling:532/540`, `modLicense:636`).

---

### FM-0133 — `modPoruke.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | Missing tabela/BuildCache fail = placeholder `[KEY]`; ne razlikuje missing-key od store-unavailable | P1 | **Tačno** — `BuildCache` `Exit Sub` na nenađenu tabelu (`:34`), pa i `Exists`-grana (`:9`) i `EH` (`:17`) vraćaju `[KLJUC]` | P2 | Modul-flag `m_storeAvailable=False` u `BuildCache`; u `Poruka` logovati unavailable jednom/sesiji | S |
| 2 | Lookup case-sensitive (`CompareMode=0`) | P1 | **Tačno** ali ključevi su compile-time literali (`Poruka("...")` ↔ `UpsertRow "..."`), nema korisničkog case-a → typo se hvata u testu | P3 | `CompareMode=1` (vbTextCompare) u `BuildCache` i `existing` — usput hvata case-duplikate | S |
| 3 | Duplicate ključevi tiho prepisani (poslednji red pobeđuje) | P1 | **Tačno** (`mCache(k)=` `:40`; `existing(k)=` `:52`) ali dup nastaje samo ručnim kvarenjem skrivene tabele | P3 | Uz vbTextCompare dodati dev-scan u `EnsurePoruke` (log ako se ključ javi >1×) | S |
| 4 | `UpsertPoruke` prepisuje administratorske izmene | P1 | **Dizajnersko ograničenje** — katalog-u-kodu je izvor istine (CLAUDE.md), `UpsertRow :287-290` namerno vraća canonical; tabela je very-hidden, nije override-sloj | Prihvaćeno | Dokumentovati „generated registry, edit kroz `UpsertPoruke`" (1 komentar) | S |
| 5 | Schema positional (`Range(1)/Range(2)`), ne named | P1 | **Tačno** po doktrini (`:39-40,51-52,288-296` čita pozicijski iako imenovane kolone postoje) — realni drift nizak (app drži fiksne 2 kol) | P3 | `GetColumnIndex(lo, COL_POR_KLJUC/TEKST)` jednom u `BuildCache`/`UpsertRow` | S/M |
| 6 | `Poruka` guta grešku bez loga | P1 | **Tačno** za `EH` (`:16-17` nema loga; missing-key ima samo `Debug.Print :12`). FM-ova rekurzija-bojazan **neosnovana** → `LogErr` bezbedan | P2 | `LogErr "modPoruke.Poruka", kljuc` u `EH` (once-per-session guard) | S |
| 7 | Cache bez auto-detekcije izmene tabele | P1 | **Dizajnersko ograničenje** — invalidacija samo eksplicitna (`:282`); tabela very-hidden, single-writer desktop | Prihvaćeno | Bez izmene; budući editor poruka MORA zvati `InvalidateCache` (dokument.) | S |
| 8 | Modul nije višejezičan | P2 | **Kontekst-Pozitivno** — i18n nije zahtev; naziv „modPoruke" ne obećava lokalizaciju | Prihvaćeno | 1 red doc: „single-locale registry, ne i18n" | S |
| 9 | Nema placeholder/format contract-a (fragment-konkatenacija) | P2 | **Tačno** kao observacija; radi ispravno, token-format je veći refaktor bez bug-a danas | P3 | `{0}`-tokeni samo oportuno kad se poruka ionako dira; bez bulk-a | S (inkr.)/L |
| 10 | Ključevi `_2/_3/_4/_5` nesemantični | P2 | **Tačno** (`:87,89,92,94` `DOK_MSG_GRESKA_PRI_CUVANJU_2..5`) — održivost smell | P3 | Preimenovati oportuno + sinhrono ažurirati `Poruka()` pozive | M |
| 11 | Veliki generated registry u jednom modulu | P2 | **Kontekst-Pozitivno** — FM sam priznaje ok; poklapa se s CLAUDE.md (katalog živi u `UpsertPoruke`) | Prihvaćeno | CSV/JSON generator tek ako build-tooling to podrži | L |

**Bilans FM-0133:** Svi nalazi su činjenično tačni protiv koda, ali su **precenjeni** (svih 7 tagovano P1). U kontekstu — app-vlasnička very-hidden 2-kol tabela, ključevi kao compile-time literali, single-writer, „katalog-kao-kod" po doktrini — realna hitnost je P2/P3. „Glavni zaključak" (5 tačaka) je sažetak redova 1-6. Dve jeftine i stvarno vredne popravke: **#6** (log store-failure u `EH`, dokazano bez rekurzije) i **#1** (razlikuj unavailable). #4/#7/#8/#11 su rad-po-dizajnu (Prihvaćeno). Ostalo P3 oportuno.

---

### FM-0134 — `modTheme.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | `DisableField` briše korisničku vrednost | P1 | **Delimično / Kontekst-Pozitivno** — `txt.value=""` (`:406`) tačno, ali SVI pozivi su mode-switch gde stale vrednost NE SME ostati (`frmOtkup:101,310,493`, `frmDokumenta:240-437`, `frmAgrohemija:1113`, `frmBankaExportPregled:730`); nijedan ne koristi kao value-preserving lock → nema živog data-loss-a | P3 | Ne dirati postojeći; dodati `DisableFieldKeepValue` tek kad zatreba preserve-lock | S |
| 2 | `DisableCombo` briše izbor/stabilni ID | P1 | **Delimično / Kontekst-Pozitivno** — `cmb.value=""` (`:420`); pozivi `cmbParcela` (`frmOtkup:488`), `cmbOtkupBlok` (`frmDokumenta:354`) čiste namerno pri mode-switch; nijedan hidden-ID reuse | P3 | Isto — bez izmene ponašanja | S |
| 3 | `ApplyTheme` potpuno guta greške | P1 | **Delimično** — `On Error Resume Next` (`:56-67`) skpouiran na telo; fail-soft tema je razumna (FM priznaje), sloj je čisto kozmetički, nikad ne kvari podatke | P3 | Dev-only log iza debug-flag; nije nužno u prod | S/M |
| 4 | `StyleControls` guta grešku po kontroli | P1 | **Delimično** — per-ctrl `OER` (`:82-122`); izolacija je zapravo razuman dizajn (jedna loša kontrola ne ruši temu), fali samo dev-log | P3 | Dev-log `FormName/ControlName/TypeName` u debug buildu | S/M |
| 5 | Klasifikacija dugmadi po imenu/caption heuristici | P1 | **Tačno** — „prikaži"→primary pa „Prikaži stornirane" postaje primary; „Storno"/„Deaktiviraj" van danger-seta→secondary (`:253-273`). Ali utiče SAMO na boju, ne ponašanje; mnoga dugmad ionako idu eksplicitnim `SetButton*` | P3 | Gde boja bitna, zvati eksplicitni `SetButton*` (već postoji) ili Tag-role | S/M |
| 6 | Danger/storno stil nedosledan | P1 | **Tačno** — `StyleStornoButton` krem+tamni (`:322-335`) vs `SetButtonDanger` crveni (`:359-363`); ista akcija dve boje | P3 | Uskladiti na jedan standard; ALI v2.24 „Vrati storno" undo → krem (soft/recoverable) može biti namerno — proveriti intent pre recolora | S |
| 7 | `IsDirectChild` fail-open heuristika | P2 | **Delimično** — fallback `True` (`:469,475,481`); najgore dupli styling (idempotentan, iste boje) — bezopasno | P3 | Nije nužno; dupli styling je idempotentan | S |
| 8 | Font fallback nije eksplicitan | P2 | **Dizajnersko ograničenje** — `SetFont` pod `OER` (`:457-461`); Windows-only desktop, Segoe UI dolazi s Windows-om | Prihvaćeno | Doc Windows-contract ili health-check probe fonta | S |
| 9 | Mnogi javni helperi koriste `Object` | P2 | **Delimično** — `SetButton* As Object` (`:347-365`) namerno radi i na runtime-dodatim kontrolama (`Controls.Add` obrazac); konkretan tip bi to slomio | P3 | Zadržati `Object` gde idu runtime kontrole | S |
| 10 | Modul meša styling i layout | P2 | **Tačno** ali minor (`LayoutTopKpiInternals :706`); FM priznaje da nije hitno | P3 | Ostaviti do sledeće funkc. izmene modula | M |
| 11 | Unknown `kind` tiho pada na default | P2 | **Delimično** — `Case Else` (`:608,636,702`); `kind` su literali iz koda, ne korisnički unos; typo→neutralni stil (kozmetika) | P3 | Opc. `Debug.Assert` u dev-u | S |
| 12 | Theme prepisuje namerno lokalno stilizovanje | P2 | **Delimično** — rekurzivni restyle; realno samo ako forma re-zove `ApplyTheme` posle lokalnog; standard je tema-pa-override (forme to rade); nema opt-out Tag-a | P3 | Doc „ApplyTheme prvo, pa semantic override"; opt-out Tag tek uz konkretan konflikt | S/M |

**Bilans FM-0134:** Svi nalazi kod-tačni, ali SVAKI je u **kozmetičkom theme sloju** — nijedan ne dira podatke ni poslovnu logiku. Dva „P1 data-loss" (DisableField/DisableCombo) su realni API-smell ali **ne manifestuju bug** jer svi pozivi žele clear-on-mode-switch. Error-swallowing P1-ovi su prihvatljiv fail-soft za temu (jedini jaz = dev-log). Heuristika/nedoslednost P1-ovi su isključivo vizuelni. Neto: **ništa nije stvarno P1**, realna hitnost P3. Najjeftiniji dobitak: uskladiti storno/danger standard (#6, uz proveru namere) i opc. `DisableFieldKeepValue` za budućnost.

---

### FM-0137 — `modStornoWarm.bas`

| # | Nalaz (FM) | FM težina | Opravdanost | Hitnost | Predlog | Napor |
|---|---|---|---|---|---|---|
| 1 | Caller dobija direktnu internu cache referencu | P1 | **Delimično / Kontekst-Pozitivno** — `Set ...= g_stornoDocs` (`:81,85`) tačno, ali jedini potrošač (`frmDokumenta:5111`) je **read-only** (zaseban `hits`, `5121-5131`), resetuje ref (`4998`), a elementi su **Variant nizovi** (value-copy) → nema puta ka koruptciji | P3 | Komentar „read-only" na `GetWarmStornoDocs`; plitka kopija jeftina ako se želi, nije nužna | S |
| 2 | `ScheduleStornoWarm` lažan scheduled state | P1 | **Tačno** — `OnTime` pa **bezuslovno** `m_warmScheduled=True` pod `OER` (`:51-54`); fail→veruje da je zakazano | P2 | `If Err.Number=0 Then m_warmScheduled=True` (flag tek po uspehu) | S |
| 3 | `CancelStornoWarm` lažan cancelled state | P1 | **Tačno** — bezuslovno `m_warmScheduled=False` (`:118-124`); ublaženo: `StornoWarmTick` re-gard (`:67-69`) čini pozni fire bezopasnim, cancel koristi tačan `m_nextWarmTime`+`WARM_PROC` | P2 | `If Err.Number<>0 Then LogErr` + poštene flag-semantike | S |
| 4 | Procedure string nije workbook-qualified | P1 | **Dizajnersko ograničenje** — `WARM_PROC` nekvalifikovan (`:28,53,121`), ali to je **konvencija celog projekta** (StanicaLock/Journaling/License); single-workbook desktop → ambiguity ne-scenario | Prihvaćeno/P3 | Ništa samo za ovaj modul; ako se hardeniše, onda SVI OnTime zajedno | S/M |
| 5 | Freshness zavisi od savršene invalidacije svih write puteva | P1 | **Delimično** — invalidacija samo iz `CommitTx` (jedini caller, `cls:79`); ali to je standardni write-wrapper, a kes hrani samo storno-picker listu (perf, ne integritet); worst-case stale red do sledeće TX | P2 | Auditovati da storno/recovery/import writes idu kroz `CommitTx`; gde ne — dodati `InvalidateStornoWarm` | M |
| 6 | Sinhroni `GetWarmStornoDocs` bez re-entry guarda | P1 | **Delimično** — `m_warming` se čita samo u `StornoWarmTick` (`:68`); `BuildWarm` ga POSTAVLJA (`:102`) ali ne PROVERAVA; re-entry samo kroz `DoEvents` (malo verovatno) | P3 | `If m_warming Then Exit Sub` na vrhu `BuildWarm` (owner guarda) | S |
| 7 | Ultimativni fallback vraća `Nothing` bez statusa | P1 | **Delimično/Tačno** — fallback pod `OER` (`:90-91`); prvi fail logovan (`:88`), drugi ne; caller (`frmDokumenta:5121`) tretira `Nothing` kao „Nema dokumenata" → empty≠failure na UI-ju (ali traži dvostruki fail, retko) | P3 | `LogErr` i u fallback grani (ili sentinel) | S |
| 8 | Nema metadata o cache-u | P2 | **Tačno** — dijagnostički jaz, nije defekt | P3 | Opc. `BuiltAt`/`Count` debug-accessor | S |
| 9 | 60 s hardcoded | P2 | **Dizajnersko ograničenje** — FM sam kaže ok, bez potrebe za config | Prihvaćeno | Izložiti samo u debug-statusu | S |
| 10 | Shutdown contract zavisi od spoljnog caller-a | P2 | **Netačno za konkretnu preporuku** — `Workbook_BeforeClose` (`doccls:78`) VEĆ zove `ShutdownApp`→`StopStornoWarm` (`:239`), `mIsShuttingDown` sprečava dupli. Rezidual (fatal-startup/self-update/kill) je uži realni edge | Prihvaćeno (std. put pokriven) | Ništa za BeforeClose; opc. osigurati da self-update putanja zove `StopStornoWarm` | S |
| 11 | `InvalidateStornoWarm` guta scheduling failure | P2 | **Delimično** — `OER` bez loga (`:43`), ali dirty=True garantuje ispravnost (sledeći `GetWarmStornoDocs` gradi sinhrono); swallow je namerno iz commit-putanje | P3 | Opc. debug-log schedule-fail; swallow ostaje | S |

**Bilans FM-0137:** Modul je zdrav; nalazi su lifecycle/ownership hardening, ne bugovi. Dve stvarno vredne jeftine popravke: **#2** (flag tek po uspehu) i **#6** (guard u `BuildWarm`) — obe S. #3/#7 jeftin log. **#1** je nemanifestovan princip (sole caller read-only + value-type elementi). **#4** je dosledna projektna konvencija (nije modul-specifičan defekt). **#10 je faktički netačan** — `BeforeClose→ShutdownApp→StopStornoWarm` već postoji, „dodaj u BeforeClose" je redundantno. **#5** je legitiman arhitektonski audit-item (P2). Ukupna hitnost P2/P3, nijedan stvarni P1.

---

## Zbirni zaključak audita

Sva tri modula: FM činjenice o kodu su uglavnom **tačne**, ali je **težina sistematski precenjena** (skoro sve P1). Nula nalaza dodiruje integritet podataka — `modPoruke` je centralizacija UI teksta, `modTheme` čist kozmetički sloj, `modStornoWarm` perf-kes browse liste sa sinhronim fallback-om.

**Stvarno vredne minimal-delta popravke (S, niska regresija):**
1. `modPoruke.Poruka.EH` → `LogErr` (FM-0133 #6; dokazano bez rekurzije).
2. `modStornoWarm.ScheduleStornoWarm` → `m_warmScheduled` tek po uspehu (FM-0137 #2).
3. `modStornoWarm.BuildWarm` → `If m_warming Then Exit Sub` na vrhu (FM-0137 #6).

**Rad-po-dizajnu / Prihvaćeno (ne dirati):** FM-0133 #4/#7/#8/#11, FM-0134 svi (kozmetika), FM-0137 #4/#9.

**Faktički netačan nalaz:** FM-0137 #10 (BeforeClose već pokriva StopStornoWarm kroz ShutdownApp).

**Napomena za FM-0134 #6 (storno boja):** pre bilo kakvog recolora proveriti da li je krem „soft" storno namerno — v2.24 uvodi „Vrati storno" undo, pa je storno tretiran kao oporavljiv, a ne terminalno-destruktivan.

---

## Ukupni zaključak DEO III (v142 delta, 38 fajlova / 8 blokova verifikovano protiv `origin/main` v2.24.0)

- **Tačnost dokumenta ostaje vrlo visoka:** dominantno **Tačno**, nula sistemskih opovrgnutih nalaza; korekcije su gotovo isključivo kalibracija težine (single-writer model, read-only dijagnostika, kozmetički slojevi, test-moduli van aktivnog E2E gate-a, 4-nivo banka integritet).
- **Novi čist P0 (aktivan gubitak): 0.** Najozbiljniji potvrđeni lanac (FM-0093 E2E false-green) je **latentan** jer gate nije pozvan u release proceduri → AUD-039, P1.
- **Novi P1 (verifikovani, minimal-delta):** (1) `frmAgrohemija` izlaz cena ≠ knjižena — `SaveMagacin` bez `overrideCena` (najvredniji nalaz delte, fix S); (2) `modAgrohemija` cena 0 tiho (fix S); (3) `modBrojevi.GenerateBrojPrijemnice` EH → `1/ddmmyy` duplikat (fix S); (4) `modMasterSync.GenerateBrojZbirne` row-count duplikat (fix M); (5) `modMasterSync.TryUpdateVozacID` True-na-neuspeh (fix S); (6) nevalidan datum → današnji (OTK+VOZ, fix S); (7) auto-otpremnica group-key meša vrste/cene (fix M); (8) VOZ link membership + prepis veza bez konflikt-politike (fix M); (9) poison spreadsheet na header-write fail (fix M); (10) `modIntegritet.WriteErr` false-count / `Empty`=PASS (fix M); (11) `modSledljivost`/`frmSledljivost` nepotpun trag prikazan kao kompletan (fix S–M); (12) stanica-mirror missing-shadow → orphan FK (fix M).
- **Novi P2:** `modProductionHealthCheck` SEF lista drift (`SEF_CANCELLED` nepostojeći) + parent-OK-posle-child-FAIL; shipped destruktivni test runneri bez env guard-a (AUD-039 familija); `modStornoWarm` lažni scheduled/cancelled flag; `modParse` single-separator (= AUD-007).
- **Već registrovano (referencirano, ne eskalirano):** AUD-002 (108.7), AUD-003 (pozicioni write), AUD-007 (121.3), AUD-016 (test dubleti), AUD-017/034 (startup), AUD-018 (108.40), AUD-037 (100.5 publish-guard), AUD-039 (E2E false-green — cela test-suite grupa).
- **Sistematski precenjeno:** `modMouseWheel`, `modTheme`, `frmMarza`, UI permission-bypass adapteri, banka-parseri (svi FM „P0/P1" → realno P2/P3).
