AgriX VBA Test Protokol
Verzija 6.6 — April 2026

Pre svega
Testovi se nikad ne pokreću na production workbook-u. Uvek napravi kopiju fajla pre testiranja. Test redovi koriste prefiks TST-PRO- i mogu se obrisati nakon testiranja.
Testovi mutiraju podatke u workbook-u. SEF live testovi mutiraju state na SEF serveru. Cancel i storno testovi su nepovratni na SEF strani.

Priprema
1. Napravi kopiju workbook-a
Sačuvaj kopiju pod imenom npr. AgriX_TEST_YYYYMMDD.xlsm. Otvori kopiju, ne original.
2. Proveri config vrednosti u tblSEFConfig
Otvori sheet koji sadrži tblSEFConfig i proveri:
KljučVrednost za testiranjeNapomenaSEF_BASE_URLURL SEF serveraMora počinjati sa https://SEF_API_KEYTvoj API ključSEF_ENVDEMO ili PRODSEF_DEBUG_LOGDAPrikazuje HTTP detalje u Immediate windowSEF_TEST_ALLOW_LIVEDADozvoljava live SEF poziveSEF_TEST_ALLOW_PRODDASamo ako testiš na produkcijskom SEF-uSEF_TEST_ALLOW_CANCEL_STORNODASamo za destruktivne testoveSEF_PAYMENT_DUE_DAYS15Ako nije setovano, default je 15
3. Otvori Immediate window
U VBA editoru: View → Immediate Window ili Ctrl+G. Ovde se ispisuju svi [PASS]/[FAIL]/[SKIP]/[INFO] rezultati u realnom vremenu.

Standardni redosled testiranja
Pokreći tačno ovim redosledom. Svaki korak zavisi od prethodnog.

Korak 1 — Schema provjera
vbRunBusinessFlowProSeedOnly
Šta radi: Proverava da sve tabele i kolone postoje. Kreira seed master podatke (ST-90001, KOOP-90001 itd.) ako ne postoje.
Očekivani rezultat: Sve linije [PASS]. Ako ima [FAIL] ovde, ne nastavljaj — schema je neispravna.
Gde gledaš rezultate: Sheet BUSINESS_FLOW_PRO_TEST_LOG

Korak 2 — Business flow regression
vbRunBusinessFlowProSuite
Šta radi: Kompletan lanac Otkup → Otpremnica → Zbirna → Prijemnica → Faktura. Proverava atomičnost, duplikat blokove, auto-link, cross-zbirna audit.
Očekivani rezultat: Sve linije [PASS]. Audit test na kraju (Cross-zbirna link audit) proverava i produkcijske podatke — ako ima [FAIL] tu, to znači da postoje stvarni podaci sa pogrešnim linkovima, ne samo test problem.
Gde gledaš rezultate: Sheet BUSINESS_FLOW_PRO_TEST_LOG, Immediate window

Korak 3 — SEF offline validacija
vbRunSEFOfflineSuite
ili sa konkretnom fakturom:
vbRunSEFOfflineSuite "FAK-00001"
Šta radi: Proverava SEF config, gradi DTO i UBL XML, validira payload — bez ijednog HTTP poziva ka SEF-u.
Očekivani rezultat: Sve [PASS]. Ako faktura ima DeliveryDate > InvoiceDate, taj test će biti [SKIP] — to je ispravno ponašanje, ne greška.
Gde gledaš rezultate: Sheet SEF_TEST_LOG, Immediate window

Korak 4 — Kreiranje test fakture za live testove
vbDim fid As String
fid = CreateSEFLiveDummyFaktura()
Debug.Print "FakturaID: " & fid
Šta radi: Kreira kompletnu fakturu sa današnjim datumom koja prolazi SEF date validaciju. Printa FakturaID u Immediate window.
Važno: Zapiši fid vrednost — trebaće ti u koracima 5, 6 i 8.
Očekivani rezultat: FakturaID koji počinje sa FAK-

Korak 5 — Live SEF slanje
vbRunSEFLiveSendSuite fid
Šta radi: Šalje fakturu na SEF, proverava response, upisuje submission record.
Očekivani rezultati:

[PASS] Live send + refresh completed — faktura primljena, HTTP 200
[PASS] Live send reached SEF and was rejected by SEF validation — SEF je primio ali odbio iz biznis razloga (validan ishod)
[SKIP] sa porukom o datumu — faktura ima invalid datume (pokreni korak 4 ponovo)
[FAIL] Live send technical failure — mrežni problem ili config greška

Gde gledaš detalje: Immediate window, sheet SEF_TEST_LOG, sheet koji sadrži tblSEFSubmission

Korak 6 — Refresh idempotency
vbRunSEFRefreshIdempotencySuite fid
Šta radi: Poziva refresh dva puta na istoj fakturi i proverava da state nije regredirao.
Očekivani rezultat: [PASS] Refresh twice did not break state

Korak 7 — Batch maintenance smoke
vbRunSEFBatchMaintenanceSmoke
Šta radi: Pokreće RefreshPendingOutboundInvoices_TX i RecoverAllStuckSEFSendingInvoices i proverava da ne crashuju.
Očekivani rezultat: Oba [PASS]

Korak 8 — Cancel/Storno (opciono, destruktivno)
⚠️ Ovo menja SEF state nepovratno. Pokreni samo na fakturi iz koraka 4.
vbRunSEFCancelLiveSuite fid
ili:
vbRunSEFStornoLiveSuite fid, "ST-001"
Šta radi: Šalje cancel ili storno zahtev na SEF. Traži unos potvrde pre izvršavanja.
Kada pokrećeš: Samo kada hoćeš da certifikuješ cancel/storno flow. Ne treba svaki put.
Potvrda: Sistem će tražiti da ukucaš tačno CANCEL FAK-XXXXX ili STORNO FAK-XXXXX pre izvršavanja.

Korak 9 — Cleanup test podataka
Nakon testiranja, obriši test redove iz tabela:
vbHardDeleteBusinessFlowTestRows
Sistem traži unos BRISI kao potvrdu. Briše sve redove koji sadrže TST-PRO- prefiks iz svih tabela, u ispravnom redosledu od child ka parent tabelama.
Alternativa — soft storno umesto brisanja:
vbSoftStornoBusinessFlowTestRows

Minimalni redosled za rutinsku provjeru
Kada ne menjaš SEF konfiguraciju i samo hoćeš da potvrdiš da sistem radi:
1. RunBusinessFlowProSuite
2. RunSEFOfflineSuite
3. HardDeleteBusinessFlowTestRows

Minimalni redosled pre svake sezone
1. RunBusinessFlowProSeedOnly       ← schema check
2. RunBusinessFlowProSuite          ← regression
3. RunSEFOfflineSuite               ← DTO/UBL check
4. fid = CreateSEFLiveDummyFaktura  ← kreiraj fixture
5. RunSEFLiveSendSuite fid          ← live smoke
6. RunSEFRefreshIdempotencySuite fid ← idempotency
7. RunSEFBatchMaintenanceSmoke      ← batch check
8. HardDeleteBusinessFlowTestRows   ← cleanup

Tumačenje rezultata
StatusZnačenjeAkcija[PASS]Test prošaoNišta[FAIL]Test paoPogledaj detalje u Immediate window[SKIP]Test preskočen zbog preduslovaPogledaj razlog — može biti očekivano[FATAL]Suite se srušio pre krajaKritična greška, pogledaj detalje[INFO]Informativna porukaSamo za kontekst

Česti problemi
[FAIL] Core tables and required columns exist
Schema nije ispravna. Proveri da li su sve ListObject tabele prisutne i da kolone odgovaraju onima u AR dokumentu.
[SKIP] Build DTO and UBL — DeliveryDate must not be later than InvoiceDate
Faktura koju testiraš ima prijemnicu sa kasnijim datumom od datuma fakture. Za offline test ovo je [SKIP], ne greška. Za live test — pokreni CreateSEFLiveDummyFaktura da dobiš ispravnu fixture fakturu.
[FAIL] Live send technical failure
Mrežni problem ili pogrešan SEF_BASE_URL/SEF_API_KEY. Proveri config, proveri internet konekciju.
[FAIL] Cross-zbirna link audit found N mismatch(es)
Postoje produkcijski otkup redovi koji su linked na otpremnicu sa drugačijim BrojZbirne. Ovo su stvarni podaci koji zahtevaju ručnu korekciju u frmOtkupniBlokovi.
MsgBox: Live SEF tests are blocked
SEF_TEST_ALLOW_LIVE nije setovano na DA u tblSEFConfig.
MsgBox: Production-like SEF environment detected
URL ne sadrži DEMO/TEST/SANDBOX pa se tretira kao produkcija. Postavi SEF_TEST_ALLOW_PROD = DA ako je to nameravano.

Šta NE raditi

Ne pokrećaj testove na production workbook-u
Ne pokrećaj cancel/storno na fakturu koja je stvarna poslovna faktura
Ne brišeš BUSINESS_FLOW_PRO_TEST_LOG i SEF_TEST_LOG sheet-ove — čuvaj ih kao audit trail
Ne menjaš seed ID-eve (ST-90001, KOOP-90001 itd.) — testovi su hardkodovani na njih
Ne pokrećaj HardDeleteBusinessFlowTestRows bez prethodnog pregleda šta će biti obrisano
