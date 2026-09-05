# Production runbook: Startup, AutoSave, Backup i Journal recovery

Status: **operativni runbook za incidente “Excel se srušio”, “unos je nestao”, “aplikacija se ne otvara”, “journal warning kaže moguć gubitak podataka”, “backup postoji ali ne znamo šta vratiti”.**

Aplikacija: **OtkupApp / AgriX Excel/VBA master**
Domen: **Workbook startup → table validation → backup → journal → TX commit → autosave → crash recovery**
Glavni kod: `src-vba/ThisWorkbook.doccls`, `src-vba/modMain.bas`, `src-vba/modJournaling.bas`, `src-vba/clsTransaction.cls`, `modDataAccess.AppendRow`

---

## 1. Kada korisnik kaže problem

Tipični incidenti:

* “Excel se srušio posle unosa.”
* “Unos koji sam malopre napravio nije tu posle ponovnog otvaranja.”
* “Aplikacija pri pokretanju prikazuje upozorenje o Journal folderu.”
* “Imam backup fajlove, koji je pravi?”
* “Workbook je read-only, a promene se ne čuvaju.”
* “AutoSave nije snimio poslednju transakciju.”
* “Aplikacija se ne otvara, vidi se samo Excel.”
* “StartApp izbaci grešku pri inicijalizaciji.”
* “Postoji journal CSV sa više redova nego Excel tabela.”
* “Puklo je tokom SEF slanja / banke / master sync-a.”

Prvo pravilo:

> Ne nastavljaj operacije i ne radi ručni save preko postojećeg fajla dok ne proveriš `Backup\`, `Journal\`, dnevni log i poslednji TX koji je mogao da se commit-uje.

Minimalni podaci koje operator mora da prikupi:

```text
Workbook path:
Workbook file name:
Da li je workbook ReadOnly:
Vreme crash-a:
Šta je korisnik radio pre crash-a:
Koja tabela/domen: Otkup, Novac, Fakture, SEF, Banka, MasterSync...
Da li je prikazan Journal warning:
Koje tabele su u warning-u:
Journal file name:
Journal row count:
Excel table row count:
Najnoviji backup file:
Najnoviji log file:
Da li postoji SEF_SENDING recovery slučaj:
```

---

## 2. Source of truth: gde se gleda

### 2.1. `ThisWorkbook.Workbook_Open`

Workbook open poziva:

```vb
StartApp
```

Ako `StartApp` padne, Excel se vraća vidljiv i prikazuje poruku da se pogleda log.

### 2.2. `modMain.StartApp`

Startup sekvenca:

1. `InitApp` ako aplikacija nije inicijalizovana.
2. `Application.Visible = False` pa `modUiFaze.FazaBoot` — **splash je prvo što
   operater vidi**, i stoji preko svih kapija ispod. (Sveska je već skrivena u
   `Workbook_Open`, prvom naredbom; ovde se skrivanje ponavlja jer je `StartApp`
   javan i pokreće se i ručno.)
3. `AccessGateOrQuit` — licenca / probni period.
4. `CheckForUpdateOnOpen` — self-update; na „Da" zakazuje `RunSelfUpdate` i izlazi.
5. `UpdateGateOrQuit` — min-version gate.
6. `modAuth.Login` — **prijava**, faza `LOGIN` istog prozora. Neuspela prijava
   otkriva svesku i zakazuje gašenje.
7. First-run setup (`SetupNewPC`) — **jedino mesto koje otkriva svesku i nastavlja**:
   bira foldere kroz `FileDialog`. Splash se skloni, setup odradi, sveska se
   vrati skrivena, splash se vrati.
8. `MOUSEWHEEL_SCROLL`.
9. `FazaBootSacekaj 1.2` pa `modOtkupUI.ShowOtkupUI` — **ekran**. Čekanje je donja
   granica trajanja splash-a, ne fiksna pauza: kad je start trajao duže, ne čeka
   ništa. Stari meni (`frmOtkupAPP`) je obrisan u koraku 7.
10. `BackupFileOnStart`.
11. `PurgeOldBackups`.
12. `PurgeOldJournals`.
13. `PurgeOldLogs`.
14. `LogAppStart`.
15. `RecoverAllStuckSEFSendingInvoices` best-effort.
16. `CheckJournalForRecovery`.
17. Ako journal warning postoji, prikazuje MsgBox.

**Sveska se od `v6-ui-214` otkriva na tačno tri mesta:** odbijena kapija (licenca,
verzija, prijava — te grane same zovu `Visible = True` pa gase aplikaciju),
first-run setup, i dugme „Otvori Excel" u ljusci (iza prava `OBL_OTVORI_EXCEL`).
Nigde drugde između klika na fajl i ekrana.

Važno:

> `CheckJournalForRecovery` samo upozorava. Ne radi automatski restore.

### 2.3. `Backup\` folder

Lokacija:

```text
ThisWorkbook.Path\Backup\
```

Ime backup fajla:

```text
<WorkbookBaseName>_YYYY-MM-DD_hhmm.<ext>
```

Backup se pravi na startu preko `ThisWorkbook.SaveCopyAs`.

### 2.4. `Journal\` folder

Lokacija:

```text
ThisWorkbook.Path\Journal\
```

Ime journal fajla:

```text
<tableName>_YYYY-MM-DD.csv
```

Primeri:

```text
tblOtkup_2026-05-06.csv
tblNovac_2026-05-06.csv
tblFakture_2026-05-06.csv
tblSEFSubmission_2026-05-06.csv
tblSEFEventLog_2026-05-06.csv
```

Svaki journal CSV ima:

```text
JournalTime;<kolone tabele...>
```

Journal se piše iz `AppendRow` odmah pri append operaciji. Journal write je best-effort i ne sme da blokira poslovnu operaciju.

### 2.5. Dnevni log

Proveriti log fajl / log sheet koji koristi `LogInfo`, `LogWarn`, `LogErr`.

Posebno tražiti:

```text
LogAppStart
LogAppShutdown
BackupFileOnStart
AutoSaveAfterCommit
CheckJournalForRecovery
StartApp
Workbook_Open
clsTransaction
RollbackTx
CommitTx
```

### 2.6. Excel tabele

Za tabele u journal warning-u proveriti:

```text
ListObject row count
poslednji redovi u tabeli
poslednji ID-jevi
da li ID-jevi iz journala postoje u Excel tabeli
```

---

## 3. Koji ID pratiš

Recovery nije po “vremenu” samo. Moraš pratiti konkretne ID-jeve koji su appendovani.

Po tabelama:

```text
tblOtkup             -> OtkupID
tblOtpremnica        -> OtpremnicaID
tblZbirna            -> ZbirnaID
tblPrijemnica        -> PrijemnicaID
tblFakture           -> FakturaID
tblFakturaStavke     -> StavkaID / FakturaID / PrijemnicaID
tblNovac             -> NovacID
tblBankaImport       -> BankaImportID
tblSEFSubmission     -> SEFSubmissionID
tblSEFEventLog       -> SEFEventID / FakturaID / SEFSubmissionID
tblAmbalaza          -> Ambalaza movement ID ako postoji, ili ReferenceID
```

Incident ticket minimum:

```text
Crash time:
Workbook path:
Backup candidate:
Journal files:
Table:
Journal row count:
Excel row count:
Missing IDs:
Existing IDs:
Last successful AutoSave log:
Last CommitTx / business action:
Recovery decision:
```

---

## 4. Normalan startup tok

Normalan `Workbook_Open` tok:

1. Otvara se workbook.
2. `StartApp` se pokreće.
3. Aplikacija validira core tabele.
4. Excel UI se sakriva.
5. Prikazuje se splash.
6. Kreira se backup kopija trenutnog workbook-a.
7. Brišu se backup/journal/log fajlovi stariji od 30 dana.
8. Loguje se app start.
9. Pokušava se recovery zaglavljenih SEF slanja.
10. Proverava se da li današnji journal ima više redova nego Excel tabela.
11. Ako nema upozorenja, splash nastavlja ka glavnoj formi.

---

## 5. Normalan TX / AutoSave tok

### 5.1. Transakcija

`clsTransaction` radi:

1. `BeginTx` čuva Excel state i gasi screen/events/manual calculation.
2. `AddTableSnapshot` pravi snapshot tabela koje mogu biti promenjene.
3. Business funkcija radi append/update.
4. Ako uspe, `CommitTx` vraća Excel state.
5. Posle commit-a poziva `AutoSaveAfterCommit`.
6. Ako padne, `RollbackTx` vraća snapshot tabela.

### 5.2. Journal

`AppendRow` poziva `WriteJournalRow`.

Značenje:

* journal može imati red i pre nego što workbook bude fizički snimljen;
* ako Excel crashuje posle append-a a pre save-a, journal može biti jedini trag;
* journal nije transakcioni source of truth sam po sebi, ali je recovery evidence.

### 5.3. AutoSave

`AutoSaveAfterCommit`:

* radi best-effort `ThisWorkbook.Save` posle uspešnog TX commit-a;
* ima debounce od 3 sekunde;
* preskače save ako je workbook read-only;
* preskače save ako workbook nema path;
* čuva `DisplayAlerts` state;
* failure ne propagira grešku korisniku, samo loguje.

Važno:

> TX može biti commitovan u memoriji, ali AutoSave može biti preskočen ili neuspešan. Zato journal i backup postoje.

---

## 6. Statusi / signali i značenje

### 6.1. Journal warning na startup-u

Poruka tipa:

```text
UPOZORENJE - Moguc gubitak podataka!
<table>: Journal ima X entries, Excel ima Y rows.
```

Značenje:

* današnji journal za tabelu ima više append redova nego sama Excel tabela;
* moguć crash ili rollback/file mismatch;
* ne znači automatski da sve treba reimportovati.

Akcija:

* stopirati operacije;
* uporediti ID-jeve iz journala i Excel tabele;
* obnoviti samo stvarno missing redove i samo ako downstream efekti to dozvoljavaju.

### 6.2. AutoSave skipped debounce

Značenje:

* prethodni save se desio pre manje od 3 sekunde;
* aplikacija namerno nije snimila svaki commit ako su preblizu.

Akcija:

* proveriti da li je kasniji save ipak uspeo;
* ako je crash bio odmah posle transakcije, journal može sadržati red koji nije u workbook-u.

### 6.3. Workbook read-only

Značenje:

* AutoSave se preskače;
* korisnik može raditi u memoriji, ali promene nisu bezbedno sačuvane.

Akcija:

* ne koristiti production workbook u read-only režimu;
* sačuvati kopiju pod novim imenom samo uz tehničku odluku;
* proveriti file lock / network share / OneDrive konflikt.

### 6.4. Workbook has no path

Značenje:

* workbook nije sačuvan na disku ili je otvoren iz privremenog konteksta;
* backup, journal i autosave nisu pouzdani.

Akcija:

* production rad nije dozvoljen dok workbook nema stabilan path.

### 6.5. Backup postoji, journal postoji, Excel tabela ne sadrži red

Značenje:

* potencijalni crash posle append-a a pre save-a;
* treba utvrditi da li je red samo u journal-u ili postoji u backup-u.

Akcija:

* uporediti backup, trenutni workbook i journal.

---

## 7. Standardni incident flow

### Korak 1: Zaustavi dalje operacije

Ako postoji sumnja na gubitak podataka:

1. Ne unositi nove dokumente.
2. Ne pokretati import/sync/SEF retry.
3. Ne overwrite-ovati workbook ručnim save-om dok se ne napravi kopija trenutnog stanja.
4. Napraviti forenzičku kopiju trenutnog workbook-a ako je moguće.

### Korak 2: Identifikuj vreme i domen

Zapiši:

```text
Vreme crash-a:
Zadnja radnja korisnika:
Domen: Otkup / Dokumenta / Faktura / Novac / Banka / SEF / MasterSync
Korisnik/operator:
```

### Korak 3: Proveri backup

U `Backup\` folderu nađi:

```text
najnoviji backup pre incidenta
najnoviji backup posle incidenta/startupa
```

Zapiši:

```text
Backup file:
Timestamp iz imena:
File size:
Da li se otvara:
Broj redova u relevantnim tabelama:
```

### Korak 4: Proveri journal

Za datum incidenta proveri sve relevantne CSV fajlove:

```text
Journal\tblX_YYYY-MM-DD.csv
```

Za svaku tabelu:

```text
Journal row count bez header-a:
Excel row count:
Poslednji JournalTime:
Poslednji ID-jevi:
```

### Korak 5: Uporedi ID-jeve, ne samo broj redova

Za svaki journal red proveri da li ID postoji u Excel tabeli.

Tabela:

```text
Table | ID | JournalTime | ExistsInExcel | ExistsInBackup | DownstreamExists | Decision
```

### Korak 6: Klasifikuj stanje

| Signal                                             | Kategorija                    | Sledeći korak              |
| -------------------------------------------------- | ----------------------------- | -------------------------- |
| journal nema missing ID-jeve                       | false alarm / row count drift | dokumentovati i nastaviti  |
| journal ima missing ID, nema downstream            | safe candidate za restore     | tehnički restore uz backup |
| journal ima missing header, ali downstream postoji | partial chain                 | domain runbook             |
| SEF submission/event missing                       | SEF recovery                  | SEF runbook                |
| banka PDF pomeren, tabela rollback                 | bank/file mismatch            | Banka runbook              |
| Google writeback urađen, Excel rollback            | MasterSync mismatch           | MasterSync runbook         |
| AutoSave read-only skipped                         | environment issue             | file lock/path recovery    |

### Korak 7: Doneti recovery odluku

Recovery nije automatski. Za svaki missing red odlučiti:

```text
Reimportovati iz journala?
Vratiti ceo backup?
Ručno rekonstruisati kroz business proceduru?
Stornirati downstream i ponoviti tok?
Eskalirati domain owner-u?
```

---

## 8. Recovery scenariji

### 8.1. Aplikacija se ne otvara / `Workbook_Open` fail

Postupak:

1. Excel treba da ostane vidljiv zbog error handler-a.
2. Proveriti log za `ThisWorkbook.Workbook_Open` i `StartApp`.
3. Proveriti da li nedostaju core tabele iz `ValidateAllTables`.
4. Proveriti da li workbook path postoji.
5. Ako je problem u formama/splash-u, tehnički owner otvara Excel vidljivo i pokreće dijagnostiku.
6. Ne raditi business operacije dok `StartApp` ne prođe.

### 8.2. Startup prikazuje Journal warning

Postupak:

1. Screenshot poruke.
2. Ne klikati dalje business operacije.
3. Otvoriti `Journal\` folder.
4. Identifikovati tabele iz poruke.
5. Za svaku tabelu uporediti journal ID-jeve sa Excel tabelom.
6. Ako nema missing ID-jeva, dokumentovati false alarm.
7. Ako ima missing ID-jeva, nastaviti po scenariju 8.3.

### 8.3. Journal ima red koji ne postoji u Excel tabeli

Postupak:

1. Identifikuj ID i tabelu.
2. Proveri da li downstream redovi postoje.
3. Proveri backup pre/posle incidenta.
4. Proveri da li je transakcija mogla biti rollbackovana.
5. Ako je red appendovan u okviru TX koji je kasnije rollbackovan, ne restore-ovati red samo zato što je u journal-u.
6. Ako je business TX uspeo, ali save nije, restore može biti potreban.
7. Tehnički owner radi restore uz ticket.

### 8.4. TX rollback se desio, ali journal sadrži append red

Ovo je očekivano moguće jer journal piše append best-effort odmah.

Postupak:

1. Proveri log za `RollbackTx` / error iz business funkcije.
2. Ako je TX rollbackovan, journal red ne znači da poslovni red treba vratiti.
3. Ne reimportovati bez domain provere.
4. Ako downstream nema ničega, verovatno nema recovery-ja.

### 8.5. AutoSave nije snimio poslednju uspešnu transakciju

Postupak:

1. Proveri log `AutoSaveAfterCommit`.
2. Traži:

   * `Saved after TX commit`;
   * `Skipped (debounce)`;
   * `Workbook read-only`;
   * `Workbook has no path`;
   * `LogErr AutoSaveAfterCommit`.
3. Ako je save preskočen i crash je bio odmah posle TX-a, proveriti journal.
4. Ako je workbook read-only, utvrditi ko je otvorio lock/kopiju.
5. Recovery kroz journal/backup.

### 8.6. Workbook je read-only

Postupak:

1. Ne nastavljati production unos.
2. Proveriti da li drugi korisnik drži fajl otvoren.
3. Proveriti network share / OneDrive / permissions.
4. Ako su promene već unesene u read-only sesiji, odmah exportovati relevantne tabele/journal ako postoji.
5. Tehnički owner odlučuje da li se pravi nova master kopija ili se promene ručno reimportuju.

### 8.7. Backup treba vratiti

Postupak:

1. Nikada ne zameniti production workbook bez kopije trenutnog fajla.
2. Napraviti folder incidenta:

```text
Incident_YYYY-MM-DD_hhmm\
  current_workbook_copy.xlsm
  chosen_backup.xlsm
  journal_files\
  logs\
```

3. Otvoriti backup read-only.
4. Uporediti relevantne tabele i ID-jeve.
5. Ako backup ima konzistentnije stanje, tehnički owner predlaže restore.
6. Poslovni owner potvrđuje da li su kasniji unosi izgubljeni i moraju se reimportovati.
7. Tek tada zameniti production fajl.

### 8.8. SEF crash/recovery

Startup pokušava `RecoverAllStuckSEFSendingInvoices`.

Postupak:

1. Ako je crash bio tokom SEF slanja, ne retry ručno odmah.
2. Proveri `tblFakture.SEFWorkflowState`.
3. Proveri `tblSEFSubmission` i `tblSEFEventLog`.
4. Ako je `SEF_SENDING`, preći na SEF runbook.
5. Ako journal ima SEF redove koji nisu u Excel-u, posebno proveriti da li postoji `SEFDocumentId`.

### 8.9. Banka import crash

Rizik:

* Excel TX rollback vrati `tblBankaImport`;
* PDF je već pomeren u `Processed` ili `Error`.

Postupak:

1. Proveri `tblBankaImport` i `Journal\tblBankaImport_...csv`.
2. Proveri `APP_BANKA_INBOX`, `APP_BANKA_PROCESSED`, `APP_BANKA_ERROR`.
3. Ne vraćati PDF u inbox bez provere dedupe-a.
4. Preći na Banka/NOVAC runbook.

### 8.10. MasterSync crash / Google writeback mismatch

Rizik:

* Excel append uspeo, Google writeback nije;
* ili Google writeback uspeo, Excel save nije.

Postupak:

1. Proveri journal za `tblOtkup`, `tblZbirna`, `tblOtpremnica` itd.
2. Proveri Google Sheet `SyncStatus`.
3. Ako Google kaže `Synced>Master`, a Excel red fali, ne reimportovati naslepo.
4. Preći na PWA MasterSync runbook.

---

## 9. Kako se radi ručni journal restore

Ovo radi samo tehnički owner.

### 9.1. Pre restore-a

Obavezno:

```text
[ ] Napravljena kopija trenutnog workbook-a
[ ] Sačuvani relevantni journal CSV fajlovi
[ ] Sačuvan backup kandidat
[ ] Identifikovani missing ID-jevi
[ ] Provereno da TX nije rollbackovan
[ ] Provereno da downstream ne pravi konflikt
[ ] Poslovni/domain owner odobrio restore
```

### 9.2. Restore princip

Ne radi “import celog journal fajla”. Radi samo missing ID-jeve.

Za svaki missing red:

1. Mapirati CSV kolone na trenutne table headers.
2. Proveriti da schema nije promenjena od momenta journala.
3. Proveriti da ID ne postoji.
4. Appendovati red.
5. Proveriti downstream reference.
6. Ručno pokrenuti relevantne status recalculation / validation procedure.

### 9.3. Posle restore-a

Uraditi:

```text
Validate relevant chain
Run domain-specific audit
Save workbook
Napraviti novi backup
Dokumentovati restore ID-jeve
```

---

## 10. Operativna pravila po domenima

### 10.1. Otkup / dokumentni chain

Ako restore-uješ `tblOtkup`, moraš proveriti:

```text
OtpremnicaID
BrojZbirne
Ambalaza movements
document chain consistency
```

### 10.2. Novac

Ako restore-uješ `tblNovac`, moraš pokrenuti:

```vb
Call UpdateFakturaStatus("FAK-...")
Call UpdateOtkupStatus("OTK-...")
```

za pogođene redove.

### 10.3. Faktura

Ako restore-uješ `tblFakture`, moraš proveriti:

```text
tblFakturaStavke
tblPrijemnica.Fakturisano
tblPrijemnica.FakturaID
SEFWorkflowState
SEFDocumentId
```

### 10.4. SEF

Ako restore-uješ `tblSEFSubmission` ili `tblSEFEventLog`, moraš proveriti:

```text
SEFSubmissionID
FakturaID
SEFDocumentId
HTTP response
SEFWorkflowState
```

Ne retry dok SEF stanje nije jasno.

### 10.5. Banka

Ako restore-uješ `tblBankaImport`, moraš proveriti:

```text
IzvorFajl
BankaReferenz
Obradjeno
NOV-* link
PDF lokacija
```

---

## 11. Admin/VBA komande

Dijagnostika:

```vb
' Provera journal warning-a
Debug.Print CheckJournalForRecovery()

' Napravi startup backup ručno
Call BackupFileOnStart

' Očisti stare journal fajlove
Call PurgeOldJournals

' Očisti stare backup fajlove
Call PurgeOldBackups

' Test autosave ponašanja
Debug.Print TestAutoSaveSmoke()

' Ručni save app-a
Call SaveApp

' Otvori Excel UI ako je sakriven
Call OpenExcel
```

SEF startup recovery:

```vb
Call RecoverAllStuckSEFSendingInvoices
```

Ne postoji bezbedna generička “RestoreJournal” procedura u trenutnom kodu. Restore je ručni tehnički postupak uz domain runbook.

---

## 12. Ko donosi odluku

### Operator sme sam

* napraviti screenshot journal warning-a;
* obustaviti dalje operacije;
* javiti poslednju radnju i vreme crash-a;
* otvoriti Excel UI kroz postojeću opciju ako treba;
* poslati backup/journal/log tehničkom owner-u.

### Tehnički owner odlučuje

* da li se koristi backup ili journal;
* da li se restore-uje pojedinačni missing red;
* da li je journal red nastao iz rollbackovane transakcije;
* da li se menja production workbook;
* da li se radi ručni CSV import;
* kako se rešava read-only/no-path stanje.

### Poslovni/domain owner odlučuje

* da li se kasniji unosi posle backup-a moraju ručno rekonstruisati;
* da li je dokumentni/finansijski/SEF chain validan posle restore-a;
* da li se rad nastavlja ili se radi nova poslovna korekcija.

### Niko ne sme bez odobrenja

* zameniti production workbook backup fajlom;
* importovati ceo journal CSV bez ID provere;
* brisati journal/backup fajlove tokom incidenta;
* nastaviti fakturisanje/SEF slanje posle journal warning-a bez provere;
* raditi production unos u read-only workbook-u;
* resetovati state ručnim brisanjem redova.

---

## 13. Checklist za zatvaranje incidenta

```text
[ ] Identifikovano vreme incidenta
[ ] Identifikovana poslednja poslovna operacija
[ ] Sačuvana kopija trenutnog workbook-a
[ ] Proveren Backup folder
[ ] Proveren Journal folder
[ ] Proveren dnevni log
[ ] Identifikovane tabele iz warning-a
[ ] Upoređeni journal ID-jevi sa Excel tabelama
[ ] Provereno da li je TX rollbackovan
[ ] Provereni downstream efekti
[ ] Doneta odluka: no-op / restore row / restore backup / domain recovery
[ ] Ako je restore rađen, dokumentovani svi redovi
[ ] Pokrenuta domain validacija posle restore-a
[ ] Workbook snimljen
[ ] Napravljen novi backup
[ ] Korisnik obavešten
```

---

## 14. Primeri odluke

### Primer A: Journal warning za `tblNovac`, jedan `NOV-*` fali

Zaključak: moguć crash posle finansijskog append-a.
Akcija: proveriti da li taj `NOV-*` postoji u backup-u, da li ima BIM trag i koji status fakture/otkupa je trebalo da promeni. Restore samo uz finansijski owner.

### Primer B: Journal ima `tblFakture` red, ali nema `tblFakturaStavke`

Zaključak: ne restore-ovati faktura header izolovano.
Akcija: proveriti da li je TX rollbackovan ili je crash bio između append-a. Faktura bez stavki je P0 i ne sme na SEF.

### Primer C: Journal ima SEF submission koji fali u Excel-u

Zaključak: moguć SEF neodređen status.
Akcija: SEF runbook. Prvo utvrditi `SEFDocumentId` i request/response, tek onda retry/recovery.

### Primer D: Workbook je read-only i operator je uneo podatke

Zaključak: AutoSave nije mogao da snimi.
Akcija: odmah exportovati/sačuvati trenutnu kopiju, ne zatvarati bez tehničkog owner-a.

### Primer E: Backup od jutros je dobar, ali tokom dana ima novih Google/PWA import-a

Zaključak: restore backup-a bi izgubio kasnije unose.
Akcija: uporediti Google/PWA/MasterSync i journal pre zamene workbook-a.

---

## 15. Poznate production rupe koje treba zatvoriti

1. Dodati kontrolisanu proceduru `RestoreJournalRows_TX(tableName, ids)` sa schema check-om.
2. Dodati `tblRecoveryEventLog`: ko je restore-ovao, šta, kada, zašto.
3. Dodati startup ekran koji prikazuje journal warning sa ID listom, ne samo row count.
4. Dodati automatski export incident paketa: current workbook copy, relevant journals, logs, latest backup.
5. Dodati alert kada AutoSave preskoči zbog read-only/no-path.
6. Dodati hard blokadu business operacija ako workbook nema path ili je read-only.
7. Dodati journal marker za TX begin/commit/rollback, da se razlikuje rollback journal od committed journal-a.
8. Dodati unique TX ID u journal redove.
9. Dodati dashboard “Journal vs Excel row mismatch”.
10. Dodati backup retention config u `tblConfig`.
11. Dodati OneDrive/network lock detection.
12. Dodati smoke test u startup checklist: `TestAutoSaveSmoke` pre produkcione sezone.

Do tada važi konzervativno pravilo:

> Journal je dokaz da je append pokušan, ne dokaz da je transakcija poslovno uspešno završena. Backup je snapshot fajla, ne nužno najnovije poslovno stanje. Recovery odluka mora gledati ID-jeve, TX/log i downstream posledice.
