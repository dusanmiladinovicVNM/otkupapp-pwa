Legenda:

✅ Rešeno / zatvoreno u kodu
🟡 Delimično rešeno / treba još mala dorada ili smoke
🔴 Otvoreno / još nije rešeno
⚫ Compile/source-of-truth provera
P0 — blokatori / compile / source-of-truth
ID	Stavka	Status	Komentar
P0-1	modBankaImportParserPdfToText.bas missing parser problem	✅	Rešeno — fajl je pronađen pod pravim imenom modBankaImportParserPdfToText.bas; parser dependency više nije blocker.
P0-2	GetBankaImportRowByID contract regresija	✅	Prvi patch je promenio shape, ali poslednji patch vraća stari 1x10 contract uz RequireColumnIndex, što je ispravno.
P1 — pre GO obavezno
P1-1 — modStorno fail-fast refaktor

Status: 🟡 skoro rešeno

modStorno diff je dobar: uvedeni su RequireStornoAllowed, RequireUpdateCell, rollback kad base funkcija ne uspe, monitoring success/fail, i uklonjen je MsgBox iz business sloja. To direktno rešava stari problem Count > 0 i silent UpdateCell.

Ostaje:

- vratiti Attribute VB_Name = "modStorno" ako je i dalje zakomentarisan
- compile smoke
- test: duplicate ID mora baciti grešku
- test: already stornirano mora fail bez promene podataka

Ocena stanja: 8.2/10, posle compile/smoke može u zatvoreno.

P1-2 — Banka parser: config path, unique temp, exit code

Status: 🟡 delimično rešeno

Parser diff rešava najvažnije:

- uklonjen hardcoded C:\Users\Dusan\...
- uveden ResolvePdfToTextExePath
- uveden unique temp txt path
- briše se stale temp file
- proverava se exitCode pdftotext.exe

To je veliki pomak.

Ostaje:

- popraviti GetBaseFileNameNoExt ako još uvek ne vraća vrednost
- koristiti CONFIG_KEY_PDFTOTEXT_EXE_PATH umesto literal "PDFTOTEXT_EXE_PATH"
- centralizovati relative path u modConfig, npr. APP_PDFTOTEXT_RELATIVE_EXE_PATH
P1-3 — modBankaImport fail-fast staging + file move atomicity

Status: 🟡 skoro rešeno

Prvi patch je rešio:

- SaveBankaImportRows koristi RequireColumnIndex
- AppendRow <= 0 je hard fail
- duplicate-only je odvojen od imported
- uvedeni statusi: imported, duplicate-only, parse/integrity/schema/append error

Drugi patch je rešio glavni batch problem:

- ImportOnePdfIntoBankaImport više ne pomera PDF odmah
- successful files idu u pendingMoves
- DB CommitTx ide pre ExecutePendingBankaFileMoves
- ako DB rollback, successful PDF-ovi se ne pomeraju u Processed

Ostaje:

- loš PDF trenutno ostaje u Inbox-u ako parse/integrity padne
- treba odlučiti: errorMoves posle rollback-a za parse/extract/integrity greške
- public ImportBankaInbox bez TX ne sme ostati opasan entrypoint
- novi potpis ImportOnePdfIntoBankaImport može slomiti stare pozive/testove

Preporučeno zatvaranje:

Public Sub ImportBankaInbox()
    ImportBankaInbox_TX
End Sub

I dodati pendingErrorMoves samo za PDF-level greške.

P1-4 — modBankaMapiranje exact-row guards

Status: 🟡 delimično rešeno

Dobar patch: uvodi RequireSingleRow, LinkNovacToOtkupStrict, RequireUpdateCell, fail-fast za NovacID, BankaImportID, OtkupID, FakturaID.

Dodatni follow-up je rešio dve važne regresije:

- GetBankaImportRowByID vraćen na stari 1x10 contract
- ValidateBankaImportNotProcessed opet proverava Stornirano

Ostaje da proveriš/popraviš:

- MapBankaImportAsKooperantBlockCore double increment/decrement bug

Ispravna logika treba da bude:

If Len(Trim$(novID)) = 0 Then
    Err.Raise ERR_BMAP_BASE + 40, "MapBankaImportAsKooperantBlockCore", _
        "SaveNovac nije vratio NovacID za OtkupID=" & otkupID
End If

LinkNovacToOtkupStrict novID, otkupID, "MapBankaImportAsKooperantBlockCore"

MapBankaImportAsKooperantBlockCore = MapBankaImportAsKooperantBlockCore + 1
preostaloZaRaspodelu = preostaloZaRaspodelu - iznosZaRed
P1-5 — MasterSync / document-chain exact-row guards

Status: 🔴 otvoreno

Raniji nalaz: AutoCreateOtpremniceFromPWA i slični link update-i ne smeju koristiti “ako postoji bar jedan red”. Za OtkupID, OtpremnicaID, ZbirnaID, PrijemnicaID, FakturaID treba Count = 1.

Ostaje:

- modMasterSync exact-row helper
- link update Otkup → Otpremnica
- link update Zbirna/Otkup/Otpremnica
- fail-fast kod duplicate/missing reference
P1-6 — modGoogleSyncOrchestrator: unlock failure critical

Status: 🔴 otvoreno

Ako SetPWAMasterSyncLock(False, ...) padne, final cycle ne sme ostati success.

Ostaje:

- failed unlock => SyncPWAFullCycle_Core = False
- monitoring CRITICAL event: PWA_MASTER_LOCK_RELEASE_FAIL
- UI/operator message: ručna provera PWA lock-a
P1-7 — modGoogleSheets.WriteSheetData non-atomic write

Status: 🔴 otvoreno

ClearSheet pre PUT values znači: ako clear uspe a write padne, Google tab može ostati prazan.

Ostaje:

- minimum: svaki WriteSheetData=False obara full sync
- bolje: staging tab write → verify → replace target
P1-8 — modGoogleSheets retry/backoff

Status: 🔴 otvoreno

Nema retry za 429/5xx.

Ostaje:

- centralni SendGoogleRequestWithRetry
- retry samo za 429, 500, 502, 503, 504
- ne retry za 400/401/403 osim specifičnog razloga
P1-9 — modGoogleAuth production OAuth/security

Status: 🔴 otvoreno

Ostaje:

- potvrditi RunGoogleAuthSetup na čistoj mašini
- zaštititi config/token sheet
- very hidden config sheet
- ne logovati secret/token
- dugoročno Windows Credential Manager/encrypted local store
P1-10 — modNovac.SaveNovac append hard fail

Status: 🔴 otvoreno

SaveNovac ne sme vratiti "" bez hard error-a ako AppendRow <= 0.

Ostaje:

- AppendRow <= 0 => Err.Raise
- caller TX rollback
P1-11 — modFaktura duplicate FakturaID guards

Status: 🔴 otvoreno

CreateFaktura je dobar, ali PrintFaktura, UpdateFakturaStatus i status/update putanje treba da proveravaju rows.Count = 1, ne samo Count > 0.

Ostaje:

- exact-row guard za FakturaID
- duplicate FakturaID => hard fail
P1-12 — modGeoParcele: save by ParcelaID

Status: 🔴 otvoreno

Trenutno public API radi po rowIndex, što je krhko ako se tabela sortira/filteruje.

Ostaje:

Public Sub SaveParcelGeoPointByID(ByVal parcelaID As String, _
                                  ByVal nCoord As Double, _
                                  ByVal eCoord As Double)

Unutra:

- FindRows(TBL_PARCELE, COL_PAR_ID, parcelaID)
- Count = 1
- onda internal row-based update
P1-13 — modGeoParcele: coordinate sanity bounds

Status: 🔴 otvoreno

Ostaje:

- posle ConvertUTM34ToLatLng proveriti lat/lng bounds
- odbiti očigledno pogrešne koordinate
P1-14 — modProductionHealthCheck proširenje

Status: 🔴 otvoreno

Health check je dobar, ali treba dodati nove P stavke.

Ostaje dodati checks za:

- duplicate IDs po kritičnim tabelama
- BankaImport saldo kolone
- pdftotext.exe path
- Google OAuth config bez ispisivanja tokena
- Geo lat/lng bounds
- PWA stale lock
- Banka open/error folder status
P1-15 — PWA otkupni list: uskladiti obračun sa kanonskim BRUTO modelom (VBA)

Status: 🔴 otvoreno (odloženo — trenutno se radi samo VBA)

Kanon (VBA `modPrint.FillOtkupSablon` + ARCHITECTURE_CHANGELOG/REFERENCE §5.12): `tblOtkup.Cena` je BRUTO (sadrži PDV nadoknadu). Otkupni list prikazuje NETO = `cena/(1+stopa)`, PDV nadoknadu kao posebnu stavku, a „za isplatu" = bruto = `kolicina*cena`.

PWA (`src/js/features/otkup/otkupni-list.js`): i modal (`showOtkupniList`) i `savePdfToDrive` računaju `vrednost = kolicina*cena` pa DODAJU PDV povrh (`ukupno = kolicina*cena*(1+stopa)`), tretirajući cenu kao osnovicu. Ako je PWA `record.cena` ista BRUTO vrednost kao `tblOtkup.Cena`, PWA prikazuje „za isplatu" ~stopa% (default 8%) VEĆI od zakonskog otkupnog lista.

Ostaje:

- PRVO proveriti da li je PWA `record.cena` identična BRUTO `tblOtkup.Cena` (sync put: `otkup-form fldCena` → GAS → `tblOtkup.Cena`; default unosa je config `Cena{vrsta}`).
- ako jeste BRUTO: uskladiti `otkupni-list.js` (`showOtkupniList` + `savePdfToDrive`) da prikazuju neto cenu/vrednost + PDV nadoknadu kao posebnu liniju, ukupno = bruto; ista formula kao VBA (`cenNeto = cena/(1+stopa/100)`).
- dodati klauzulu čl. 34 ZPDV u PWA (trenutno je nema; default tekst u `modDocStyle.OtkupKlauzulaDefault`).
- proveriti da PWA `liveTotal` (`kolicina*cena` u `otkup-form.js`) i Pregled/izveštaji koriste isti model, da ne nastane nova neusklađenost.

Nađeno tokom v6.28 (VBA OtkupSablon 1/3 A4). Čisto PWA izmena; VBA strana je ispravna.

P2 — hardening / održavanje
ID	Stavka	Status	Komentar
P2-1	modStammdatenSync: ukloniti hardcoded TOTAL_STAMMDATEN_TABS = 13	🔴	Broj tabova izračunati iz StammdatenTabs().
P2-2	modGoogleSheets.ParseValuesJson testovi ili zamena parsera	🔴	Dodati testove za comma, quotes, newline, UTF-8.
P2-3	GetSpreadsheetID duplicate spreadsheet warning	🔴	Ako ima više exact-name fajlova, log WARN.
P2-4	Redakcija Google HTTP response body logova	🔴	Ne logovati PII/token/payload fragmente.
P2-5	modSEFTax config-driven poreski defaulti	🔴	10, "S", "35" u config.
P2-6	clsSEFValidationResult dopuniti ili ukloniti	🔴	Trenutno shell/scaffold.
P2-7	modSEFPersistance naming cleanup	🔴	Dugoročno Persistence, ne Persistance.
P2-8	modAgrohemija fail-fast + typo cleanup	🔴	SaveMagacin append hard fail; typo Doabvljacu.
P2-9	frmOtkupAPP.btnBanka_Click busy lock	🔴	Koristiti SetSidebarEnabled False/True, refresh badge.
P2-10	ClearParcelGeo_TX briše COL_PAR_DATUM_GEO	🔴	Trenutno ostaje stari geo timestamp.
P2-11	GoogleParcelaExists retry	🔴	Kratak retry posle sync-a.
P2-12	APP_VERSION alignment	🔴	2.2.1 uskladiti sa stvarnim release markerom.
P2-13	modConfig PDFTOTEXT constants cleanup	🟡	Dodata konstanta je dobar smer, ali treba koristiti konstantu svuda i centralizovati relative path.
P3 — cleanup / kvalitet / test coverage
ID	Stavka	Status	Komentar
P3-1	SEF XML snapshot testovi	🔴	Ručno sklapanje UBL XML-a treba snapshot test.
P3-2	Shared modDataAccessGuards	🔴	Izvući RequireSingleRow, RequireActiveSingleRow, IsStorniranoValue, RequireAppendRow.
P3-3	modGoogleAuth.GetAccessToken formatiranje	🔴	Samo čitljivost.
P3-4	modGeoParcele geo source parametar	🔴	Umesto hardcoded "selenium".
P3-5	Ukloniti “PATCH” komentare iz production modula	🟡	Dobri su za review, ali dugoročno treba normalan module header.
