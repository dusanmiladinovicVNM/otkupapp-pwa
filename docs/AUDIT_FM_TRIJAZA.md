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
