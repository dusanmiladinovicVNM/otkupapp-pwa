# Production runbook: SEF slanje faktura

Status: **operativni runbook za incident „Ne mogu da pošaljem fakturu.”**
Aplikacija: **OtkupApp**
Glavni tok: Excel/VBA SEF modul, `frmSEF`, `tblFakture`, `tblSEFSubmission`, `tblSEFEventLog`

---

## 1. Kada korisnik kaže: „Ne mogu da pošaljem fakturu”

Cilj nije odmah kliknuti retry. Cilj je prvo utvrditi:

1. koja je faktura u pitanju;
2. da li postoji lokalni submission;
3. da li je HTTP poziv već otišao;
4. da li je SEF vratio `SEFDocumentId`;
5. da li je problem validacioni, tehnički ili poslovno-pravni;
6. da li je bezbedno retry-ovati bez pravljenja duplog dokumenta.

Minimalni podaci koje operator mora da prikupi od korisnika:

* `BrojFakture`, ako ga zna;
* `FakturaID`, ako ga vidi u aplikaciji;
* kupac;
* vreme pokušaja slanja;
* screenshot/status poruke iz SEF forme;
* da li je korisnik već kliknuo slanje više puta.

---

## 2. Source of truth: gde se gleda

### 2.1. Prvo mesto: SEF forma

Otvoriti `frmSEF` / ekran **SEF upravljanje**.

Na izabranoj fakturi proveriti:

* `FakturaID`
* `BrojFakture`
* `Kupac`
* `SEFWorkflowState`
* `SEFStatus`
* `SEFDocumentId`
* `SEFVersionNo`
* poslednju grešku
* donju tabelu **SEF Event Log**

Forma učitava događaje preko `GetSEFEventsForFaktura(fakturaID)` i prikazuje vreme, tip, poruku i detalje događaja.

### 2.2. Glavne tabele

#### `tblFakture`

Primarni red za status fakture. Proveriti:

| Kolona                | Značenje                                                             |
| --------------------- | -------------------------------------------------------------------- |
| `FakturaID`           | primarni interni ID fakture                                          |
| `BrojFakture`         | poslovni broj fakture / broj dokumenta                               |
| `KupacID`             | kupac za fakturu                                                     |
| `SEFWorkflowState`    | interni workflow state                                               |
| `SEFStatus`           | poslednji eksterni status koji je vratio SEF                         |
| `SEFDocumentId`       | eksterni ID dokumenta na SEF-u                                       |
| `SEFLastErrorCode`    | poslednji tehnički/SEF error code                                    |
| `SEFLastErrorMessage` | poslednja poruka greške                                              |
| `SEFPayloadHash`      | hash UBL payload-a koji je slat                                      |
| `SEFSubmissionIDLast` | poslednji interni submission/request ID                              |
| `SEFVersionNo`        | verzija SEF pokušaja                                                 |
| `PoslatNaSEF`         | `Da` tek kada lokalni pipeline dođe do `SEF_SENT` ili `SEF_ACCEPTED` |
| `SEFSentAt`           | vreme prvog uspešnog lokalnog slanja                                 |
| `SEFLastSyncAt`       | vreme poslednjeg sync-a sa SEF statusom                              |

#### `tblSEFSubmission`

Audit red za konkretan pokušaj slanja. Proveriti:

| Kolona                  | Značenje                                                              |
| ----------------------- | --------------------------------------------------------------------- |
| `SEFSubmissionID`       | interni submission ID; koristi se i kao `requestId` u SEF POST pozivu |
| `FakturaID`             | faktura kojoj pripada pokušaj                                         |
| `VersionNo`             | verzija slanja                                                        |
| `WorkflowStateAtSubmit` | state u trenutku kreiranja submission-a                               |
| `CreatedAt`             | vreme lokalnog kreiranja submission-a                                 |
| `SubmittedAt`           | vreme snimanja rezultata slanja                                       |
| `FinishedAt`            | vreme završetka pokušaja                                              |
| `SubmissionStatus`      | `CREATED`, `SENT`, `ACCEPTED`, `REJECTED`, `FAILED`                   |
| `PayloadHash`           | hash tačnog XML payload-a                                             |
| `RequestFormat`         | očekivano `XML`                                                       |
| `RequestBody`           | UBL XML koji je poslat ili koji se ponovo koristi za retry            |
| `ResponseBody`          | raw response od SEF-a                                                 |
| `HttpStatus`            | HTTP status                                                           |
| `ApiStatus`             | status iz SEF API odgovora                                            |
| `CorrelationId`         | correlation/global ID ako SEF vrati                                   |
| `SEFDocumentId`         | eksterni SEF dokument ID ako je dobijen                               |
| `ErrorCode`             | error code                                                            |
| `ErrorMessage`          | error message                                                         |
| `OperatorName`          | Windows/Excel korisnik                                                |

#### `tblSEFEventLog`

Timeline incidenta. Proveriti:

| Kolona            | Značenje                                          |
| ----------------- | ------------------------------------------------- |
| `SEFEventID`      | ID log događaja                                   |
| `FakturaID`       | faktura                                           |
| `SEFSubmissionID` | pokušaj slanja                                    |
| `EventTime`       | vreme događaja                                    |
| `EventType`       | tip događaja                                      |
| `Message`         | ljudski čitljiva poruka                           |
| `Details`         | hash, request ID, SEFDocumentId, HTTP/API detalji |
| `OperatorName`    | operator                                          |

Tipični `EventType`: `STATE_CHANGED`, `HTTP_SENT`, `HTTP_RESPONSE`, `VALIDATION_FAILED`, `SYNC_OK`, `SYNC_FAILED`, `SEF_ACCEPTED`.

### 2.3. Journal i backup

Za crash/recovery proveriti foldere uz workbook:

* `Journal\tblSEFSubmission_YYYY-MM-DD.csv`
* `Journal\tblSEFEventLog_YYYY-MM-DD.csv`
* `Journal\tblFakture_YYYY-MM-DD.csv`, ako je bilo append operacija
* `Backup\*.xls*`

Journal je zaštita od Excel crash-a i piše se na `AppendRow`. Nije zamena za `tblSEFSubmission` i `tblSEFEventLog`, nego recovery signal.

### 2.4. Debug log

`SEF_DEBUG_LOG=DA` u `tblSEFConfig` uključuje `Debug.Print` HTTP response log u VBA Immediate Window.

To **nije production audit log** jer nije trajno sačuvan. Koristi se samo za live dijagnostiku dok je sesija otvorena. Za incident source of truth ostaju `tblSEFSubmission` i `tblSEFEventLog`.

---

## 3. Koji ID transakcije pratiš

Uvek prati ove ID-jeve zajedno:

1. `FakturaID` — interni poslovni dokument. Primer: `FAK-00008`.
2. `BrojFakture` — broj fakture koji vidi računovodstvo/kupac.
3. `SEFSubmissionIDLast` / `SEFSubmissionID` — interni submission ID, npr. `SFS-00012`. Ovo je **requestId** koji se šalje SEF-u u URL-u `sales-invoice/ubl?requestId=...`.
4. `SEFDocumentId` — eksterni ID dokumenta u SEF-u. Ako postoji, SEF zna za dokument.
5. `SEFPayloadHash` / `PayloadHash` — dokaz da li retry koristi isti payload ili je kreiran novi.
6. `CorrelationId` — ako ga SEF vrati, čuva se na submission-u.

Incident ticket mora imati minimum:

```text
FakturaID:
BrojFakture:
KupacID / Kupac:
SEFWorkflowState:
SEFStatus:
SEFSubmissionIDLast:
SEFDocumentId:
SEFPayloadHash:
HttpStatus:
ApiStatus:
ErrorCode:
ErrorMessage:
OperatorName:
Vreme pokušaja:
```

---

## 4. Kako znaš da li je dokument poslat SEF-u

### Dokument je poslat / SEF ga zna ako:

* `tblFakture.SEFDocumentId` nije prazan; ili
* poslednji red u `tblSEFSubmission` za fakturu ima `SubmissionStatus = SENT` ili `ACCEPTED` i `SEFDocumentId` nije prazan; ili
* `SEFStatus` je `SENT`, `NEW`, `DRAFT`, `ACCEPTED`, `REJECTED`, `CANCELLED` ili `STORNO` uz postojeći `SEFDocumentId`.

Tada **ne šalji novu fakturu**. Radi refresh/status, cancel ili storno kroz propisani tok.

### Dokument verovatno nije poslat ako:

* nema `SEFDocumentId`;
* poslednji `SubmissionStatus` je `FAILED` ili `CREATED`;
* `SEFWorkflowState = SEF_TECH_FAILED`;
* event log ima `SYNC_FAILED` ili tehničku grešku, bez kasnijeg uspešnog `HTTP_RESPONSE` sa `SEFDocumentId`.

Tada retry može biti dozvoljen samo po pravilima iz sekcije 6.

### Dokument je u neodređenom stanju ako:

* `SEFWorkflowState = SEF_SENDING`; ili
* postoji `HTTP_SENT`, ali nema snimljenog `HTTP_RESPONSE`; ili
* aplikacija/Excel je pukla između HTTP poziva i TX2 snimanja rezultata; ili
* postoji kontradikcija između `tblFakture` i `tblSEFSubmission`.

U neodređenom stanju **ne pravi novi submission ručno**. Prvo uradi recovery iz sekcije 7.

---

## 5. Stanja i značenje

### Interni workflow: `SEFWorkflowState`

| State             | Značenje                                   | Operator sme                                     |
| ----------------- | ------------------------------------------ | ------------------------------------------------ |
| `LOCAL_DRAFT`     | faktura još nije finalizovana              | ne šalje se                                      |
| `LOCAL_FINALIZED` | faktura je finalna lokalno                 | može slanje                                      |
| `SEF_READY`       | spremna za SEF                             | može slanje                                      |
| `SEF_SENDING`     | slanje je započeto                         | ne retry; prvo recovery                          |
| `SEF_SENT`        | lokalni submit uspeo, postoji SEF dokument | refresh status                                   |
| `SEF_ACCEPTED`    | SEF prihvatio                              | ne retry; samo storno ako poslovno odobreno      |
| `SEF_REJECTED`    | SEF odbio                                  | ne retry unchanged; korekcija + resubmit flow    |
| `SEF_TECH_FAILED` | tehnički pad slanja                        | retry dozvoljen ako se reuse-uje isti submission |
| `SEF_SYNC_ERROR`  | refresh statusa pao                        | ponoviti refresh                                 |
| `SEF_STORNO`      | stornirano                                 | finalno, ne slati ponovo bez pravne odluke       |

### Eksterni status: `SEFStatus`

`SEFStatus` je poslednji status koji vraća SEF API. Ne mora biti isti kao interni workflow.

Primeri validnih kombinacija:

* `SEFWorkflowState = SEF_SENT`, `SEFStatus = SENT`
* `SEFWorkflowState = SEF_SENT`, `SEFStatus = DRAFT`
* `SEFWorkflowState = SEF_ACCEPTED`, `SEFStatus = ACCEPTED`
* `SEFWorkflowState = SEF_REJECTED`, `SEFStatus = REJECTED`
* `SEFWorkflowState = SEF_SENT`, `SEFStatus = STORNO` ili `CANCELLED` posle refresh-a

---

## 6. Da li smeš da retry-uješ

### Dozvoljen retry

Retry je dozvoljen samo ako je:

* `SEFWorkflowState = SEF_TECH_FAILED`; i
* poslednji `SubmissionStatus` je `FAILED` ili `CREATED`; i
* nema uspešnog submission-a (`SENT` ili `ACCEPTED`) za istu fakturu; i
* nema `SEFDocumentId` koji ukazuje da SEF već zna za dokument.

U tom slučaju `SendInvoiceToSEF_TX(fakturaID)` radi kontrolisani retry:

* uzima poslednji `SEFSubmissionID`;
* uzima postojeći `RequestBody`;
* uzima postojeći `PayloadHash`;
* ne kreira novi payload;
* ne kreira novi submission;
* šalje isti `requestId` prema SEF-u.

Ovo je glavni mehanizam protiv duplog dokumenta kod tehničkog retry-ja.

### Nije dozvoljen retry

Ne retry-ovati direktno ako je stanje:

* `SEF_SENDING` — prvo recovery;
* `SEF_SENT` — prvo refresh status;
* `SEF_ACCEPTED` — finalno, nema retry;
* `SEF_REJECTED` — potrebna poslovna korekcija i resubmit flow;
* `SEF_STORNO` — finalno/pravni slučaj;
* postoji bilo koji uspešan submission `SENT` ili `ACCEPTED`.

### Rate limit / HTTP greške

Ako je `HttpStatus = 429`, sistem postavlja `ApiStatus = RATE_LIMITED`, `ErrorCode = 429`, a `ErrorMessage` može sadržati `Retry-After` vrednost.

Postupak:

1. ne klikati više puta;
2. sačekati prema `Retry-After`, ako postoji;
3. proveriti da li postoji `SEFDocumentId`;
4. ako je završilo u `SEF_TECH_FAILED`, retry je dozvoljen samo kao reuse istog submission-a.

---

## 7. Recovery procedure

### 7.1. Faktura je zaglavljena u `SEF_SENDING`

Simptom:

* korisnik kaže da slanje „vrti” ili je Excel pukao;
* `SEFWorkflowState = SEF_SENDING`;
* dugme `Recover Sending` je aktivno u SEF formi.

Postupak:

1. Otvori SEF formu.
2. Izaberi fakturu.
3. Zabeleži `FakturaID`, `SEFSubmissionIDLast`, `SEFDocumentId`, `SEFPayloadHash`.
4. Ako postoji `SEFDocumentId`, klikni **Osveži status** ili pokreni `RefreshSEFStatus_TX(fakturaID)`.
5. Ako nema `SEFDocumentId`, klikni **Recover Sending** ili pokreni `RecoverStuckSEFSendingInvoice(fakturaID)`.
6. Recovery bez `SEFDocumentId` prebacuje fakturu u `SEF_TECH_FAILED` i čuva isti `submissionID`.
7. Tek posle toga retry kroz **Retry slanje na SEF**.

Admin VBA alternativa:

```vb
Call RecoverStuckSEFSendingInvoice("FAK-00008")
```

Za masovni recovery:

```vb
Call RecoverAllStuckSEFSendingInvoices
```

### 7.2. Faktura je u `SEF_SYNC_ERROR`

Simptom:

* slanje je ranije uspelo;
* postoji `SEFDocumentId`;
* refresh statusa je pao.

Postupak:

1. Ne šalji ponovo.
2. Klikni **Osveži status** ili pokreni:

```vb
Call RefreshSEFStatus_TX("FAK-00008")
```

Ako novi refresh vrati finalni status, sistem vraća workflow u konzistentno stanje kroz dozvoljenu tranziciju.

### 7.3. SEF odbio fakturu: `SEF_REJECTED`

Simptom:

* `SEFWorkflowState = SEF_REJECTED`;
* `SEFStatus = REJECTED`;
* `SEFLastErrorMessage` ili submission `ErrorMessage` sadrži razlog.

Postupak:

1. Ne retry-ovati istu fakturu bez izmene.
2. Poslovni vlasnik/računovodstvo analizira grešku.
3. Ispraviti podatke: kupac, PIB, stavke, iznos, broj fakture ili drugo što SEF traži.
4. Nakon odobrenja kliknuti **Pripremi za ponovno slanje** ili pokrenuti:

```vb
Call PrepareRejectedInvoiceForResubmit("FAK-00008")
```

5. Ova procedura postavlja fakturu u `SEF_READY`, čisti `SEFSubmissionIDLast` i pravi novi tok slanja.
6. Tek zatim ponovo poslati.

### 7.4. Postoji `SEFDocumentId`, ali `PoslatNaSEF = Ne`

Ovo je nekonzistentno lokalno stanje.

Postupak:

1. Ne šalji ponovo.
2. Zabeleži sve ID-jeve.
3. Pokreni refresh:

```vb
Call RefreshSEFStatus_TX("FAK-00008")
```

4. Ako refresh uspe, sistem treba da upiše status i očisti grešku.
5. Ako refresh ne uspe, slučaj ide tehničkom owner-u.

### 7.5. `tblSEFSubmission` kaže `SENT`/`ACCEPTED`, ali `tblFakture` nije ažuran

Postupak:

1. Ne šalji ponovo.
2. Proveri poslednji submission za `SEFDocumentId`.
3. Ako postoji `SEFDocumentId`, radi `RefreshSEFStatus_TX(fakturaID)`.
4. Ako ne postoji `SEFDocumentId`, ali `ResponseBody` ga sadrži, tehnički owner mora ručno analizirati response pre bilo kakvog slanja.
5. Ne uređivati ručno ćelije bez backup-a i bez zapisa u incident ticket-u.

### 7.6. Excel crash / gubitak podataka

Postupak:

1. Ne slati ponovo pre poređenja.
2. Proveriti `Journal\` fajlove za isti datum.
3. Porediti broj redova u journal-u sa Excel tabelama.
4. Proveriti `Backup\` kopije.
5. Ako journal ima više redova nego workbook, vratiti podatke iz journal/backup kopije pre nastavka operacije.

---

## 8. Kako sprečavamo dupli dokument

Sistem ima više zaštita:

1. `ValidateFakturaForSEF` blokira slanje ako faktura već ima uspešan submission (`SENT` ili `ACCEPTED`).
2. `SEFWorkflowState` state machine ne dozvoljava proizvoljne tranzicije.
3. Retry iz `SEF_TECH_FAILED` koristi isti `SEFSubmissionID`, isti `RequestBody` i isti `PayloadHash`.
4. `SEFSubmissionID` se šalje SEF-u kao `requestId`, što omogućava idempotentni retry na nivou zahteva.
5. `SEFPayloadHash` i `PayloadHash` služe za proveru da nije promenjen XML između pokušaja.
6. `SEFDocumentId` je hard stop: ako postoji, ne pravi se novi submission; radi se refresh/cancel/storno.
7. `PoslatNaSEF` se postavlja na `Da` tek kada workflow dođe do `SEF_SENT` ili `SEF_ACCEPTED`.

Operativno pravilo:

> Ako vidiš `SEFDocumentId`, ne klikći slanje. Ako vidiš `SEF_SENDING`, ne klikći retry. Ako vidiš `SEF_TECH_FAILED`, retry je dozvoljen samo reuse-om istog submission-a.

---

## 9. Standardni incident flow

### Korak 1: Identifikuj fakturu

U SEF formi pronađi fakturu po `FakturaID` ili `BrojFakture`.

Zapiši:

```text
FakturaID=
BrojFakture=
KupacID=
Workflow=
SEFStatus=
SEFDocumentId=
LastSubmissionID=
PayloadHash=
LastError=
```

### Korak 2: Učitaj event log

U donjoj tabeli SEF forme pročitaj poslednje događaje.

Traži posebno:

* poslednji `HTTP_SENT`;
* poslednji `HTTP_RESPONSE`;
* `SYNC_FAILED`;
* `VALIDATION_FAILED`;
* `SEF_ACCEPTED`;
* promenu state-a.

### Korak 3: Učitaj submission

U `tblSEFSubmission` filtriraj po `FakturaID`.

Sortiraj po `CreatedAt` opadajuće i gledaj poslednji red.

Zapiši:

```text
SEFSubmissionID=
SubmissionStatus=
HttpStatus=
ApiStatus=
SEFDocumentId=
ErrorCode=
ErrorMessage=
PayloadHash=
CreatedAt=
SubmittedAt=
FinishedAt=
```

### Korak 4: Klasifikuj problem

| Signal                                                 | Kategorija     | Sledeći korak                                      |
| ------------------------------------------------------ | -------------- | -------------------------------------------------- |
| `ERR_SEF_CONFIG`, missing `SEF_BASE_URL`/`SEF_API_KEY` | konfiguracija  | tehnički owner popravlja `tblSEFConfig`            |
| missing kupac, PIB, stavke, broj fakture, iznos        | validacija     | poslovni owner ispravlja podatke                   |
| `HTTP_ERROR`, HTTP 0, timeout                          | tehnički       | ako je `SEF_TECH_FAILED`, retry istog submission-a |
| HTTP 429                                               | rate limit     | čekati `Retry-After`, zatim kontrolisani retry     |
| HTTP 400/409/422, `REJECTED`                           | SEF validacija | poslovna korekcija + resubmit flow                 |
| `SEF_SENDING`                                          | neodređeno     | recovery, nikad direktan retry                     |
| `SEF_SENT` bez finalnog statusa                        | pending        | refresh status                                     |
| `SEF_ACCEPTED`                                         | finalno        | nema retry                                         |

### Korak 5: Izvrši jedinu dozvoljenu akciju

| State             | Dozvoljena akcija                                        |
| ----------------- | -------------------------------------------------------- |
| `LOCAL_FINALIZED` | Pošalji na SEF                                           |
| `SEF_READY`       | Pošalji na SEF                                           |
| `SEF_TECH_FAILED` | Retry slanje na SEF, isti submission                     |
| `SEF_SENDING`     | Recover Sending                                          |
| `SEF_SENT`        | Osveži status                                            |
| `SEF_SYNC_ERROR`  | Osveži status                                            |
| `SEF_REJECTED`    | Pripremi za ponovno slanje samo posle poslovne korekcije |
| `SEF_ACCEPTED`    | bez retry-ja; samo poslovno odobren storno ako treba     |

### Korak 6: Posle akcije proveri konzistentnost

Posle svake akcije proveri:

```text
tblFakture.SEFWorkflowState
tblFakture.SEFStatus
tblFakture.SEFDocumentId
tblFakture.SEFLastErrorCode
tblFakture.SEFLastErrorMessage
tblFakture.SEFSubmissionIDLast
tblFakture.SEFPayloadHash
tblSEFSubmission.SubmissionStatus
tblSEFEventLog poslednji događaj
```

Incident se može zatvoriti tek kada je jedno od sledećeg tačno:

* faktura je `SEF_SENT`/`SEF_ACCEPTED` i ima `SEFDocumentId`;
* faktura je `SEF_REJECTED`, razlog je dokumentovan, i poslovni owner je preuzeo korekciju;
* faktura je vraćena u `SEF_READY` posle odobrene korekcije;
* faktura je `SEF_TECH_FAILED`, ali je incident eskaliran tehničkom owner-u sa svim ID-jevima;
* faktura je stornirana/cancelled uz poslovno-pravnu odluku.

---

## 10. Operaterske akcije kroz UI

U `frmSEF` postoje sledeće akcije:

| Dugme                        | Kada se koristi                                   | Šta radi                                                         |
| ---------------------------- | ------------------------------------------------- | ---------------------------------------------------------------- |
| `Pošalji na SEF`             | `LOCAL_FINALIZED` ili `SEF_READY`                 | poziva `SendInvoiceToSEF_TX`                                     |
| `Retry slanje na SEF`        | `SEF_TECH_FAILED`                                 | poziva `SendInvoiceToSEF_TX`, ali reuse-uje prethodni submission |
| `Osveži status`              | `SEF_SENT` ili `SEF_SYNC_ERROR`                   | poziva `RefreshSEFStatus_TX`                                     |
| `Pripremi za ponovno slanje` | `SEF_REJECTED`                                    | poziva `PrepareRejectedInvoiceForResubmit`                       |
| `Cancel`                     | eksterni status `DRAFT`, `NEW` ili `ERROR`        | poziva `CancelInvoiceOnSEF_TX`                                   |
| `Storno`                     | eksterni status `SENT`, `ACCEPTED` ili `REJECTED` | poziva `StornoInvoiceOnSEF_TX`                                   |
| `Recover Sending`            | `SEF_SENDING`                                     | poziva `RecoverStuckSEFSendingInvoice`                           |
| `Refresh Pending`            | masovni refresh pending faktura                   | poziva `RefreshPendingOutboundInvoices_TX`                       |
| `Recover All Sending`        | masovni recovery zaglavljenih sending faktura     | poziva `RecoverAllStuckSEFSendingInvoices`                       |

---

## 11. Admin/VBA komande

Koristiti samo ako UI nije dovoljan ili ako tehnički owner radi incident.

```vb
' Slanje / kontrolisani retry
Debug.Print SendInvoiceToSEF_TX("FAK-00008")

' Refresh statusa po SEFDocumentId
Call RefreshSEFStatus_TX("FAK-00008")

' Recovery jedne zaglavljene fakture
Call RecoverStuckSEFSendingInvoice("FAK-00008")

' Recovery svih zaglavljenih faktura
Call RecoverAllStuckSEFSendingInvoices

' Refresh svih pending outbound faktura
Call RefreshPendingOutboundInvoices_TX

' Priprema rejected fakture za novi submission nakon korekcije
Call PrepareRejectedInvoiceForResubmit("FAK-00008")

' Cancel na SEF-u
Debug.Print CancelInvoiceOnSEF_TX("FAK-00008", "Odobreno od računovodstva, razlog...")

' Storno na SEF-u
Debug.Print StornoInvoiceOnSEF_TX("FAK-00008", "Odobreno od računovodstva, razlog...", "STORNO-...")
```

Ne koristiti ručno editovanje state kolona osim kao poslednju meru i samo uz backup, ticket i odobrenje tehničkog owner-a.

---

## 12. Poslovno-pravne odluke

### Tehnički owner sme sam da odluči

* retry `SEF_TECH_FAILED` kada nema `SEFDocumentId` i reuse-uje se isti submission;
* refresh `SEF_SENT` / `SEF_SYNC_ERROR`;
* recovery `SEF_SENDING` u `SEF_TECH_FAILED` kada nema `SEFDocumentId`;
* ispravka konfiguracije `tblSEFConfig`.

### Poslovni owner / računovodstvo odlučuje

* ispravka podataka posle `SEF_REJECTED`;
* da li se rejected faktura šalje ponovo;
* korekcija kupca, PIB-a, iznosa, stavki, datuma ili broja fakture;
* da li se kupcu šalje obaveštenje.

### Pravni/računovodstveni owner mora da odobri

* cancel dokumenta na SEF-u;
* storno dokumenta na SEF-u;
* bilo kakav slučaj gde postoji `SEFDocumentId`, a korisnik traži „pošalji ponovo”;
* mogućnost duplog dokumenta;
* ručnu intervenciju nad već poslatim dokumentom.

### Niko ne sme bez odobrenja

* ručno brisati `SEFDocumentId`;
* ručno čistiti `SEFSubmissionIDLast` osim kroz `PrepareRejectedInvoiceForResubmit`;
* menjati `SEFWorkflowState` mimo state machine-a;
* slati novu fakturu ako postoji uspešan submission;
* retry-ovati `SEF_SENDING` bez recovery-ja.

---

## 13. Checklist za zatvaranje incidenta

Pre zatvaranja ticket-a upisati:

```text
[ ] FakturaID identifikovan
[ ] BrojFakture identifikovan
[ ] Proveren tblFakture
[ ] Proveren tblSEFSubmission
[ ] Proveren tblSEFEventLog
[ ] Utvrđen SEFSubmissionID / requestId
[ ] Utvrđeno da li postoji SEFDocumentId
[ ] Utvrđeno da li je retry dozvoljen
[ ] Ako je retry rađen, potvrđeno da je reuse-ovan isti submission/payload
[ ] Ako je rejected, dokumentovana poslovna korekcija
[ ] Ako je cancel/storno, postoji poslovno-pravno odobrenje
[ ] Posle akcije proveren finalni workflow/status
[ ] Korisnik obavešten
```

---

## 14. Primeri odluke

### Primer A: `SEF_TECH_FAILED`, nema `SEFDocumentId`

Zaključak: tehnički pad.
Akcija: retry dozvoljen, ali samo postojeći submission/request body/payload hash.
Ne kreirati novi submission ručno.

### Primer B: `SEF_SENDING`, nema response-a

Zaključak: neodređeno.
Akcija: `RecoverStuckSEFSendingInvoice`. Ako nema `SEFDocumentId`, sistem prebacuje u `SEF_TECH_FAILED`, pa retry koristi isti submission.

### Primer C: `SEF_SENT`, `SEFStatus = DRAFT`

Zaključak: SEF zna za dokument, ali eksterni status još nije finalan.
Akcija: refresh statusa. Ne slati ponovo.

### Primer D: `SEF_REJECTED`, error kaže da PIB nije validan

Zaključak: poslovno/validaciona greška.
Akcija: računovodstvo ispravlja kupca/PIB, zatim `PrepareRejectedInvoiceForResubmit`, pa novo slanje.

### Primer E: Postoji `SEFDocumentId`, korisnik kaže „pošalji opet”

Zaključak: visok rizik duplog dokumenta.
Akcija: zabranjeno ponovno slanje. Refresh/status, zatim poslovno-pravna odluka za cancel/storno ako treba.

---

## 15. Poznate production rupe koje treba zatvoriti

Ovaj runbook pokriva trenutni kod, ali za jači production readiness treba dodati:

1. Persistovan HTTP debug log umesto oslanjanja na Immediate Window.
2. Poseban `tblOpsDecisionLog` za poslovno-pravna odobrenja: ko, kada, šta je odobrio, razlog.
3. Automatski job/scheduler za `RefreshPendingOutboundInvoices_TX`, ako aplikacija treba da radi bez ručnog refresh-a.
4. Eksplicitnu proveru u SEF portalu po `BrojFakture` kada je stanje `SEF_SENDING` bez `SEFDocumentId`, pre retry-ja u kritičnim slučajevima.
5. Dokumentovanu potvrdu SEF idempotency garancije za `requestId`.
6. Admin ekran/filter za `tblSEFSubmission` po `FakturaID`, da se incident ne rešava direktnim kopanjem po tabelama.

Do tada, pravilo je konzervativno: **ako postoji šansa da je SEF kreirao dokument, ne šalji novu fakturu dok se status ne potvrdi.**
