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

P1-16 — Link-write storno-guard: novac → otkup/faktura (LinkNovacToOtkupStrict + ApplyAvans*)

Status: 🔴 otvoreno

Nađeno tokom audita korupcije OtpremnicaID (stale mActiveOtpID → stornirana otpremnica; fix commit-i ac0fb57 + c8198ab, grana claude/malina-purchase-data-xujz8g). Srodne rupe ISTE klase („upiši referencu na dokument bez provere da cilj nije storniran"), ali PARAMETARSKE (vrednost dolazi kao argument, nije stale modularna promenljiva) → niža verovatnoća od potvrđenog HIGH buga. Dopuna za P1-4: LinkNovacToOtkupStrict već ima Count=1 exact-row guard, ali NE i storno-proveru cilja.

1) modBankaMapiranje.LinkNovacToOtkupStrict (:2126; pozivi MapBankaImportAsKooperant :433, MapBankaImportAsKooperantBlockCore :864)
   - Ima RequireSingleRow (postoji + jedinstven) za NovacID i OtkupID, ali NEMA storno-proveru otkupa.
   - Rizik: uplata čija „poziv na broj“ pogodi STORNIRAN otkup veže se za mrtav dokument; potom uplata nestaje iz open-balance matcha (saldo isključuje stornirane) → novac „izgubljen“.
   - Fix: posle RequireSingleRow TBL_OTKUP, proveriti da je LookupValue(TBL_OTKUP, COL_OTK_ID, otkupID, COL_STORNIRANO) <> "Da"; ako je storniran → NE vezuj (tiho preskoči + WARN log da operater ručno reši; ne raise, da ne obori ceo bank-import TX).

2) modNovac.ApplyAvansToOtkup (:1139, COL_NOV_OTKUP_ID) i ApplyAvansToFaktura (:591, COL_NOV_FAKTURA_ID)
   - Nema storno-proveru cilja (otkup/faktura). Delimično kompenzovano: pri stornu se veza čisti (modStorno.ResetNovacOtkupLink :1156 → COL_NOV_OTKUP_ID = "").
   - Calleri obično prosleđuju svež otkupID (SaveOtkupMulti_TX) ili otvoren blok (frmBankaExportPregled „Primeni avans na blok“) → niži rizik, ali „Primeni avans na blok“ može da gađa stariji otkup.
   - Fix: pre vezivanja proveriti da cilj (otkupID/fakturaID) nije storniran; ako jeste → preskoči + poruka operateru.

Ostaje:
- odluka skip-vs-raise po putanji (bank-import TX = skip + WARN; ručni „Primeni avans“ = poruka + preskoči, ne obarati akciju).
- eventualno deliti helper sa P3-2 (RequireActiveSingleRow / IsStorniranoValue).
- smoke: uplata/avans ka storniranom otkupu ne sme da se veže; posle storna otkupa veza ostaje očišćena.

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
P3-6	Izgubljeni blokovi: „Preuzmi obe klase po BrDok“	🔴	Adopt radi po jednom OtkupID; Klasa I+II (isti BrDok, 2 OtkupID) se preuzimaju posebno.
P3-7	Izgubljeni blokovi: auto-predlog ciljne otpremnice	🔴	Detekcija ne zna „pravu“ novu otpremnicu; operater bira ručno. Predložiti aktivnu sa istim BrojOtpremnice/BrojZbirne.
P3-8	Izgubljeni blokovi: zaglavlja kolona u lost-modu	🔴	„stara otp“ se prikazuje u koloni „Vrednost“ (kozmetika); dati zasebno zaglavlje/kolonu.

P3-6..8 — Izgubljeni blokovi (dorade; prioritet za slobodno vreme)

Kontekst: feature „Izgubljeni blokovi“ (commit-i cfa12e4 + fix 315bd04). Komponente:
- modDokumenta.GetLostOtkupBlokovi (detekcija, inference),
- modDokumenta.ReassignOtkupToOtpremnica_TX (bezbedan re-point = opcija A: menja samo OtpremnicaID + BrojZbirne, čuva OtkupID),
- modHelpers.CheckVerwaisteDokumente (dashboard sekcija #5),
- modOtkupBlok (panel: ToggleLostMode / LoadLostBlokovi / AdoptSelectedLostBlok).
Radi i potvrđeno u produkciji.

P3-6 — „Preuzmi obe klase po BrDok“
AdoptSelectedLostBlok i ReassignOtkupToOtpremnica_TX rade po jednom OtkupID. Otkup sa Klasa I i II ima dva reda (isti BrojDokumenta, dva OtkupID) -> oba se pojave kao izgubljena i preuzimaju se posebno.
Dorada: preuzmi sve redove istog BrojDokumenta odjednom (obrazac StornoOtkupByBrDok_TX). Konkretno: dodati ReassignOtkupByBrDok_TX(brDok, targetOtpID) koja FindRows po COL_OTK_BR_DOK i re-pointuje svaki NEstornirani red; ili petlja u adopt-u. Paziti: ne dirati već stornirane redove.

P3-7 — auto-predlog ciljne otpremnice
GetLostOtkupBlokovi ne zna koja je „prava“ nova otpremnica; operater ručno bira cilj u levoj listi (mActiveOtpID) pre Preuzmi. Čest scenario je storno + ponovni unos sa ISTIM poslovnim brojem (BrojOtpremnice/BrojZbirne).
Dorada (opciono): kad postoji aktivna otpremnica sa istim BrojOtpremnice (ili BrojZbirne) kao „stara“, predložiti/preselektovati je kao cilj. Smanjuje rizik preuzimanja na pogrešnu otpremnicu.

P3-8 — zaglavlja kolona u lost-modu
U „Izgubljeni“ režimu lista BLOKOVI (mLstBlok) koristi normalna zaglavlja (BLOK_CAPS), a „stara otp: X“ se prikazuje u koloni sa zaglavljem „Vrednost“ -> zbunjujuće.
Dorada: privremeno zameniti zaglavlja u lost-modu (npr. kolona „Stara otpremnica“) ili zaseban mali grid. Čista kozmetika.

P3-9	Palete re-point: auto-aneks (višak) / deficit gajbica	🔴	ReassignPaleteToPrijemnica_TX sad samo UPOZORI kad se broj gajbica po klasi razlikuje (re-point + KG-sync rade samo na poklapanju). Dorada: auto-aneks paleta DATIRANA NA ORIGINAL za višak (ne preko GetOrCreateOpenPaleta — da ne padne na današnju paletu); kontrolisan deficit.

P3-9 — Palete re-point: višak/deficit gajbica (warn-only za sada)
Kontekst: P1 — modPaletniList.ReassignPaleteToPrijemnica_TX + UI „Palete“ mod u recovery panelu (frmDokumenta). Re-point + KG-sync (skaliranje neta) rade kad se broj gajbica PO KLASI poklapa. Kad se razlikuje → samo upozorenje u statusnoj liniji; operater rešava ručno.
Razlog odlaganja (svesno): auto-rutiranje viška preko GetOrCreateOpenPaleta stavilo bi ga na DANAŠNJU otvorenu paletu → pogrešan datum/lokacija za robu od pre par dana.
Dorada (kad bude vremena): za višak praviti NAMENSKU aneks paletu datiranu na original; za deficit kontrolisano skidanje sa poslednje delimične stavke uz potvrdu. Edge koji „ne bi trebalo da se dešava“ → nizak prioritet.

---

## Storno/ispravka framework (PR: `claude/central-storno-framework-j1ku1h`)

P1-STORNO-1 — Proći / review + merge PR „Centralni storno/ispravka framework" (PR #130)
Status: 🟡 spremno za merge (0 konflikata sa `main`, `RunStornoTestSuite` = FAIL=0 / **22 testa**, statika čista), čeka review + ručni smoke pa merge + release.
Sadrži i **redizajn DUPLI vs PONIŠTENJE** (DUPLI = razveži/odvezuje → deca prežive; PONIŠTENJE = kaskadni storno celog nizvodnog toka), **ownership pravilo** (otpremnica obara zbirnu samo ako je jedina), **hladnjača/eksterni gate** (`kupac == MALINA_DEFAULT_KUPAC`; eksterni kupac → prijemnica netaknuta), **palete kroz motor** (`DetachOsirocenePaletaStavke_TX` → header/reopen/storno-prazne, co-tenant netaknut), i **pre-merge review fixeve** (simple-path nulling → STORNO; atomarnost DUPLI zbirne u jednoj TX + guard; block-free guard → MANUAL; revers preview ne duplira dvojni upis).
Pre merge (ručni smoke onoga što NIJE ručno probano — unit-testovi jesu):
- Zbirna ISPRAVKA (zbirna sa prijemnicom) → snimi novu → potvrdi „zamena? DA" → proveri: otpremnice+prijemnica+paletne stavke na novoj, stara odvezana/stornirana.
- Malina 1:1: PONIŠTENJE otpremnice/zbirne → ceo lanac stornian + **prazna paleta stornirana** + deljena paleta umanjena/otvorena; DUPLI → otpremnice u „čeka zbirnu", palete osirоćene.
- Normalni mod, hladnjača kupac, 2 otpremnice na jednoj zbirni: PONIŠTENJE **jedne** → druga+zbirna+prijemnica žive; PONIŠTENJE **zbirne** → sve stornirano.
- Eksterni kupac: PONIŠTENJE zbirne → **prijemnica NETAKNUTA**.
- Ponovo `RunStornoTestSuite` posle importa (mora ostati FAIL=0).
Posle merge-a: `tools/release.sh <verzija>` (planiran `vba-v2.19.0`) → Excel `ImportAllVBA → Compile → snimi → ship → Fleet`; `RELEASE_NOTES.md` je već pripremljen.

P2-STORNO-2 — Recovery pregled pending/failed ispravki (poseban panel)
Status: 🟡 delimično. Pending/FAILED/MANUAL_REQUIRED kontekst je vidljiv u `tblStornoVeze` + Monitor (`Monitor_Event`) → zadovoljava zahtev („Monitor ili poseban pregled"). ALI panel „Osiroćeni dokumenti" ga još NE lista (prikazuje osirotele prijemnice/zbirne), a poruke ga pominju. Dorada: ili dopuniti recovery panel da čita `modStornoContext.GetPendingCorrections()`, ili doterati tekst poruka (da ne upućuju na panel koji ih ne prikazuje). `CountPendingRecovery()` već postoji za badge.

P2-STORNO-3 — Audit: „nuliranje" dokumenta umesto storna (isti bag širom koda)
Status: 🟡 delimično. **Storno flow REŠEN** (PR #130): sve putanje (`modStornoFlow` DUPLI/PONIŠTENJE/ISPRAVKA + **simple** storno otpremnice i zbirne) sada idu kroz `RecalcOrStornoEmptyZbirna_TX` → prazna zbirna se STORNIRA, ne ostaje aktivna 0/0. Ostaje audit OSTATKA koda. Originalni bag: dokument koji treba da nestane ostane AKTIVAN sa 0/0 (tihi mismatch — aktivan „prazan" dokument). Zadatak: pretražiti SVUDA gde se KG/ambalaža/gajbice postavljaju na 0 a red ostaje aktivan (nije stornirano), pa odlučiti treba li storno.
Ispravan obrazac (referenca): `modPaletniList.bas:1903-1907` — PRVO `COL_STORNIRANO="Da"` PA onda vrednosti na 0.
Kandidati/polazna tačka (grep `RequireUpdateCell/UpdateCell ... , 0` + `Recalc/Rekalk/nulir`):
- `modDokumentInvariant.RecalculateZbirnaFromOtpremnice_TX` / `ApplyKlasaRecalc` — cela prazna zbirna: sada je `modStornoFlow` stornira (OK). OSTAJE: PER-KLASA red (npr. Klasa II → 0 dok Klasa I živi) ostaje aktivan 0 — proveriti da li je legitimno (klasa stvarno 0) ili treba čišćenje.
- `modPaletniList.bas:1905-1907` — OK (storno + 0), samo potvrditi.
- Proći i: `modOtkupBlok`, `modAmbalaza`, `frmDokumenta`, `modDokumenta` (bilo koji reset KG/amb na 0 bez storna).
Rezultat: lista mesta + odluka (storno vs legitiman 0) + po potrebi popravka; idealno regres test po nalazu.

P2-STORNO-4 — Jednokratni repair alati za POSTOJEĆE pokvarene podatke (nastale pre framework-a)
Status: 🔴 otvoreno (opciono, NE blokira merge — novi kod sprečava dalje kvarenje; ovi alati čiste zatečeno stanje).
Kontekst: podaci pokvareni RANIJIM ručnim stornom / bagovitim run-om pre ovog framework-a ostaju u zatečenom stanju (novi kod ih NE popravlja retroaktivno).
1) `FixNuliraneZbirne` — prođi `tblZbirna`; svaku AKTIVNU zbirnu bez ijedne aktivne otpremnice (0/0) → `StornoZbirna_TX`. Referenca iz produkcije: ZBR-00643/644/645 (aktivne 0,00 dok prijemnica pokazuje na njih).
2) `FixStalePaletaHeaders` — prođi `tblPaleta`; preračunaj header (gajbe/neto/amb + Bruto, Status) iz AKTIVNIH `tblPaletaStavka`; paleta bez ijedne aktivne stavke → storno; ispod kapaciteta → reopen. Referenca: run pre palete-fixa `67a4381` (PAL-00150 header 240 a stvarno 206 aktivnih + Zatvorena; PAL-00151 prazna a Otvorena). Reuse paletni motor gde može (`DetachOsirocenePaletaStavke_TX` je per-prijemnica; ovde treba PER-PALETA self-heal — proveriti postoji li već helper `RecomputePaleta*` pre pisanja novog).
Zajednički zahtevi: idempotentno + transakciono + izveštaj (koliko dirano) + pokreće se svesno (Alt+F8), best-effort (ne rušiti na jednom lošem redu). Ako operater potvrdi da treba — dodati u ovoj ili sledećoj grani.
