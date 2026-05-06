1. Monitoring infrastruktura

Postavljen je end-to-end monitoring tok:

VBA → GAS Web App → monitorPublic → Monitoring.gs → OtkupApp_Monitoring_PROD

Monitoring fajl:

OtkupApp_Monitoring_PROD

Korišćeni monitoring sheet-ovi:

Events
Errors
Health
SEFStatus
Backups
Alerts
AuditCritical

VBA šalje evente preko MONITORING_ENDPOINT, tj. GAS Web App URL-a:

https://script.google.com/macros/s/.../exec

Secret model:

GAS Script Property: MONITORING_INGEST_SECRET
Excel tblSEFConfig key: MONITORING_SECRET
2. VBA monitoring config

Monitoring config se čita iz:

tblSEFConfig

sa kolonama:

ConfigKey
ConfigValue

Korišćeni config key-evi:

MONITORING_ENDPOINT
MONITORING_SECRET
MONITORING_ENV

Dodato je safe čitanje config-a:

SafeConfigValue(...)

i robustan reader koji traži tblSEFConfig kroz workbook i čita ConfigKey / ConfigValue.

APP_VERSION se čita iz modConfig konstante:

APP_VERSION

DeviceId se šalje kao lokalni identitet računara/korisnika, npr:

DESKTOP-JHQJCQH|Dusan
3. Security cleanup u modMonitoring

Uvedeno je da debug više ne printuje ceo JSON body.

Umesto:

Body: {... monitoringSecret ...}

debug prikazuje samo:

Body length
Body preview redacted
eventType present
HTTP Status
HTTP Response

Dodate su sanitizacione funkcije:

Monitoring_SanitizeText
Monitoring_SanitizePayloadJson

Sanitizuju se:

MONITORING_SECRET
GOOGLE_ACCESS_TOKEN
GOOGLE_REFRESH_TOKEN
SEF_API_KEY

Payload i poruke se skraćuju ako su predugi.

Timeout model:

HTTP_TIMEOUT_MS = 1200
HTTP_DEBUG_TIMEOUT_MS = 10000

Production monitoring je best-effort i ne sme da zamrzava Excel.

4. Monitoring test suite

Napravljen je test modul:

modMonitoringTests

Sa testovima:

TestMonitoring_All
TestMonitoring_Config
TestMonitoring_HTTP
TestMonitoring_ErrorEvent
TestMonitoring_SEFUnknown
TestMonitoring_BackupSuccess
TestMonitoring_BackupFail

Potvrđen je HTTP ingest:

HTTP Status: 200
HTTP Response: {"success":true,...}
Monitor_Test result: True

Test suite pokriva slanje u:

Events
Errors
SEFStatus
Backups
Alerts
Health
AuditCritical
5. ThisWorkbook.Workbook_Open

Uveden monitoring na otvaranje aplikacije:

VBA_APP_OPEN

U EH bloku se šalje:

Monitor_Error

sa:

moduleName = ThisWorkbook
procedureName = Workbook_Open
entityType = App
entityId = Startup
correlationId = VBA-STARTUP

Monitoring je best-effort i ne blokira startup.

6. modMain.StartApp

Uveden lifecycle monitoring za startup aplikacije:

VBA_STARTAPP_START
VBA_STARTAPP_SUCCESS

U EH bloku:

Monitor_Error

sa:

moduleName = modMain
procedureName = StartApp
entityType = App
entityId = Startup
correlationId = VBA-STARTUP

Za journal recovery warning uveden event:

JOURNAL_RECOVERY_WARN

StartApp ostaje orchestration layer; detaljni monitoring za SEF/backup/druge tokove ide u njihove module.

7. SEF startup recovery — modSEFService

Za proceduru:

RecoverAllStuckSEFSendingInvoices

uvedeni event-i:

SEF_STARTUP_RECOVERY_START
SEF_RECOVERY_INVOICE_FOUND
SEF_RECOVERY_INVOICE_SUCCESS
SEF_RECOVERY_INVOICE_FAIL
SEF_STARTUP_RECOVERY_SUCCESS
SEF_STARTUP_RECOVERY_FAIL

Uz Monitor_Error za:

pojedinačnu fakturu
globalni recovery failure

Correlation/entity princip:

entityType = Faktura / SEF
entityId = fakturaID / StartupRecovery
correlationId = fakturaID / SEF-STARTUP-RECOVERY
8. SEF send flow — modSEFService.SendInvoiceToSEF_TX

Uveden monitoring za slanje fakture na SEF.

Event posle lokalnog TX1 commita:

SEF_SEND_START

Finalni event posle TX2 commita:

SEF_SEND_ACCEPTED
SEF_SEND_SUCCESS
SEF_SEND_REJECTED
SEF_SEND_FAIL

Kritični event za slučaj kada je lokalno stanje već WF_SEF_SENDING:

SEF_SEND_EXCEPTION_AFTER_LOCAL_SENDING

Monitoring payload koristi:

fakturaID
submissionID
response.apiStatus
response.httpStatus
response.sefDocumentId
response.errorCode
response.errorMessage

Ne šalje se:

UBL XML
rawBody
API key
token
full payload
9. SEF status refresh — modSEFStatusSync

Za:

RefreshSEFStatus_TX

uvedeni event-i:

SEF_STATUS_ACCEPTED
SEF_STATUS_REJECTED
SEF_STATUS_PENDING
SEF_STATUS_TERMINAL
SEF_STATUS_UPDATE
SEF_STATUS_REFRESH_FAIL
SEF_STATUS_REFRESH_EXCEPTION

Za:

RefreshPendingOutboundInvoices_TX

uvedeni event-i:

SEF_REFRESH_PENDING_START
SEF_PENDING_REFRESH_INVOICE_FAIL
SEF_REFRESH_PENDING_SUMMARY
SEF_REFRESH_PENDING_FAIL

Summary prati:

scannedCount
refreshedCount
skippedTerminalCount
failedCount
10. Fakture — modFaktura

Za:

CreateFaktura_TX

uvedeno:

FAKTURA_CREATE_SUCCESS
FAKTURA_CREATE_FAIL
Monitor_Error

Monitoring se vezuje za:

entityType = Faktura
entityId = FakturaID
correlationId = FakturaID

Pokrivena TX granica koja dira:

tblFakture
tblFakturaStavke
tblPrijemnica
tblNovac
11. Novac — modNovac

Za:

SaveNovac_TX

uvedeno:

NOVAC_SAVE_SUCCESS
NOVAC_SAVE_FAIL
Monitor_Error

Correlation ID se bira po prioritetu:

fakturaID
otkupID
brojDok
partnerID

Za:

ApplyAvansToFaktura_TX

uveden samo fail monitoring:

AVANS_APPLY_TO_FAKTURA_FAIL
Monitor_Error

Bez success eventa, da se ne pravi šum.

12. Otkup — modOtkup

Za:

SaveOtkup_TX

uvedeno:

OTKUP_SAVE_SUCCESS
OTKUP_SAVE_FAIL
Monitor_Error

Za:

SaveOtkupMulti_TX

uvedeno:

OTKUP_MULTI_SAVE_SUCCESS
OTKUP_MULTI_SAVE_FAIL
Monitor_Error

Pokriveni podaci:

kooperantID
stanicaID
vrstaVoca
kolicina
brDok
resultI
resultII
hasKlasaII
13. Dokumentni lanac — modDokumenta

Za dokumentni tok:

Otkup → Otpremnica → Zbirna → Prijemnica → Faktura

uveden je fail-only monitoring.

Pokrivene TX procedure:

SaveOtpremnica_TX
SaveOtpremnicaMulti_TX
SaveZbirna_TX
SaveZbirnaMulti_TX
SavePrijemnica_TX
SavePrijemnicaMulti_TX

Event:

DOKUMENT_SAVE_FAIL

plus:

Monitor_Error

Bez success eventa za dokumente, da monitoring ne bude preglasan.

14. Banka mapiranje — modBankaMapiranje

Pokriven modul koji mapira:

tblBankaImport → tblNovac

Dodati helperi:

Monitor_BankaMapSuccess
Monitor_BankaMapFail

Pokriveni TX wrapper-i:

AutoMapBankaImportRow_TX
MapBankaImportAsKupac_TX
MapBankaImportAsKooperant_TX
MapBankaImportAsOM_TX
MapBankaImportAsKooperantBlock_TX
MapBankaImportAsKooperantBlockManual_TX
SkipBankaImportRow_TX
AutoMapAllBankaImport_TX

Event-i:

BANKA_MAP_SUCCESS
BANKA_MAP_FAIL
BANKA_IMPORT_SKIP
BANKA_AUTOMAP_ALL_START
BANKA_AUTOMAP_ALL_SUMMARY
BANKA_AUTOMAP_ALL_FAIL

Za batch automap prati se:

OpenRows
Mapped
NotMapped
MappedBeforeFail

Za pojedinačna mapiranja šalje se:

bankaImportID
resultId
partnerType
partnerId
linkedEntityId
15. Backup monitoring

Uveden/predviđen monitoring za backup testove:

BACKUP_SUCCESS
BACKUP_FAIL

Kroz testove se šalje u:

Events
Backups
Errors
Alerts
Health

Backup ostaje non-blocking za startup.

16. Finalni event set po oblastima
Core app
VBA_APP_OPEN
VBA_STARTAPP_START
VBA_STARTAPP_SUCCESS
JOURNAL_RECOVERY_WARN
SEF
SEF_STARTUP_RECOVERY_START
SEF_RECOVERY_INVOICE_FOUND
SEF_RECOVERY_INVOICE_SUCCESS
SEF_RECOVERY_INVOICE_FAIL
SEF_STARTUP_RECOVERY_SUCCESS
SEF_STARTUP_RECOVERY_FAIL
SEF_SEND_START
SEF_SEND_ACCEPTED
SEF_SEND_SUCCESS
SEF_SEND_REJECTED
SEF_SEND_FAIL
SEF_SEND_EXCEPTION_AFTER_LOCAL_SENDING
SEF_STATUS_ACCEPTED
SEF_STATUS_REJECTED
SEF_STATUS_PENDING
SEF_STATUS_TERMINAL
SEF_STATUS_UPDATE
SEF_STATUS_REFRESH_FAIL
SEF_STATUS_REFRESH_EXCEPTION
SEF_REFRESH_PENDING_START
SEF_PENDING_REFRESH_INVOICE_FAIL
SEF_REFRESH_PENDING_SUMMARY
SEF_REFRESH_PENDING_FAIL
Fakture
FAKTURA_CREATE_SUCCESS
FAKTURA_CREATE_FAIL
Novac
NOVAC_SAVE_SUCCESS
NOVAC_SAVE_FAIL
AVANS_APPLY_TO_FAKTURA_FAIL
Otkup
OTKUP_SAVE_SUCCESS
OTKUP_SAVE_FAIL
OTKUP_MULTI_SAVE_SUCCESS
OTKUP_MULTI_SAVE_FAIL
Dokumenti
DOKUMENT_SAVE_FAIL
Banka
BANKA_MAP_SUCCESS
BANKA_MAP_FAIL
BANKA_IMPORT_SKIP
BANKA_AUTOMAP_ALL_START
BANKA_AUTOMAP_ALL_SUMMARY
BANKA_AUTOMAP_ALL_FAIL
Backup
BACKUP_SUCCESS
BACKUP_FAIL
17. Pokriveni production rizici

Monitoring sada pokriva:

pokretanje aplikacije
startup lifecycle
journal warning
SEF slanje
SEF recovery
SEF status refresh
kreiranje fakture
unos novca
primenu avansa na fakturu
unos otkupa
multi-klasa otkup
dokumentni lanac failove
bankarsko mapiranje
skip bankarske stavke
batch automap banke
backup success/fail
centralne VBA error-e
health update
alert generation
critical audit za kritične događaje

18. Status

Monitoring infrastruktura je potvrđena end-to-end.

Core VBA monitoring modul je podešen.

Config čitanje radi.

HTTP ingest radi.

Test suite radi.

Glavne production TX granice su pokrivene.

SEF observability je najdetaljnije pokriven.

Banka mapiranje je dodato kao poslednja velika production-critical oblast.
19. GAS Web App endpoint

Monitoring ingest ide preko jednog GAS Web App endpoint-a:

https://script.google.com/macros/s/.../exec

VBA ne gađa poseban URL za code.gs ili monitoring.gs, nego deployment URL celog Apps Script projekta.

VBA šalje POST JSON na taj endpoint.

Glavna akcija u request-u:

{
  "action": "monitorPublic"
}
20. GAS public ingest contract

VBA šalje standardizovan JSON payload sa poljima:

action
monitoringSecret
environment
source
severity
eventType
userId
role
deviceId
appVersion
module
functionName
entityType
entityId
correlationId
message
payload

Primer potvrđenog VBA body-ja:

{
  "action": "monitorPublic",
  "environment": "PROD/DEV",
  "source": "VBA",
  "severity": "INFO",
  "eventType": "VBA_MONITORING_TEST",
  "deviceId": "DESKTOP-JHQJCQH|Dusan",
  "appVersion": "2.2.1",
  "module": "modMonitoring",
  "functionName": "Monitor_Test",
  "correlationId": "VBA-MONITORING-TEST"
}

Secret se šalje kao:

monitoringSecret

ali se više ne printuje u VBA debug output-u.

21. GAS secret validation

Na GAS strani se koristi Script Property:

MONITORING_INGEST_SECRET

Na VBA/Excel strani ekvivalent je:

tblSEFConfig → MONITORING_SECRET

Efektivno pravilo:

VBA MONITORING_SECRET value == GAS MONITORING_INGEST_SECRET value

Ako se vrednosti poklapaju, GAS prihvata event.

Ako se ne poklapaju, GAS vraća JSON odgovor sa neuspehom, ali HTTP konekcija i dalje radi.

22. GAS response contract

Uspešan ingest vraća:

{
  "success": true,
  "eventId": "...",
  "timestamp": "...",
  "severity": "INFO",
  "component": "VBA Client"
}

To je potvrđeno end-to-end testom:

HTTP Status: 200
HTTP Response: {"success":true,...,"component":"VBA Client"}
Monitor_Test result: True

Za VBA monitoring validacija se smatra uspešnom samo ako postoje oba uslova:

HTTP 2xx
response sadrži "success": true
23. GAS routing po tipu eventa

GAS prima sve evente kroz isti public ingest, a zatim ih razvrstava u monitoring sheet-ove.

Centralni ulaz:

monitorPublic

Efektivno ponašanje:

svaki validan event → Events
ERROR/CRITICAL event → Errors
SEF event → SEFStatus
backup event → Backups
alert-worthy event → Alerts
critical audit event → AuditCritical
component state → Health
24. GAS Events log

Events je osnovna tabela za sve monitoring događaje.

U nju idu event-i kao:

VBA_APP_OPEN
VBA_STARTAPP_START
VBA_STARTAPP_SUCCESS
FAKTURA_CREATE_SUCCESS
NOVAC_SAVE_SUCCESS
OTKUP_SAVE_SUCCESS
BANKA_MAP_SUCCESS
SEF_SEND_START
SEF_STATUS_UPDATE
BACKUP_SUCCESS

Glavna svrha:

centralni vremenski log svega što se desilo u sistemu
25. GAS Errors log

Errors prima greške iz VBA i kritične poslovne failove.

U nju idu događaji preko:

Monitor_Error

i događaji sa severity:

ERROR
CRITICAL

Tipični izvori:

Workbook_Open
StartApp
SendInvoiceToSEF_TX
RefreshSEFStatus_TX
CreateFaktura_TX
SaveNovac_TX
SaveOtkup_TX
modDokumenta TX failovi
modBankaMapiranje TX failovi

Polja koja se šalju:

moduleName
procedureName
entityType
entityId
correlationId
errorNumber
errorDescription
errorSource
26. GAS Health table

Health služi za trenutni status komponente.

Ažurira se kroz monitoring evente po komponentama:

VBA Client
SEF
Backup
Banka
App Startup

Efektivno ponašanje:

poslednji signal po komponenti
poslednja severity vrednost
poslednji eventType
poslednji timestamp

Ovo omogućava brz pregled:

da li VBA klijent šalje evente
da li je SEF flow zdrav
da li backup radi
da li banka mapiranje puca
27. GAS SEFStatus table

SEFStatus prima specijalizovane SEF evente.

Pokriveni SEF event-i:

SEF_STARTUP_RECOVERY_START
SEF_RECOVERY_INVOICE_FOUND
SEF_RECOVERY_INVOICE_SUCCESS
SEF_RECOVERY_INVOICE_FAIL
SEF_SEND_START
SEF_SEND_ACCEPTED
SEF_SEND_SUCCESS
SEF_SEND_REJECTED
SEF_SEND_FAIL
SEF_SEND_EXCEPTION_AFTER_LOCAL_SENDING
SEF_STATUS_ACCEPTED
SEF_STATUS_REJECTED
SEF_STATUS_PENDING
SEF_STATUS_TERMINAL
SEF_STATUS_UPDATE
SEF_STATUS_REFRESH_FAIL
SEF_STATUS_REFRESH_EXCEPTION
SEF_REFRESH_PENDING_SUMMARY

Glavna polja:

invoiceLocalId
businessInvoiceNo
sefStatus
localStatus
sefRequestId
sefInvoiceId
attemptCount
lastHttpCode
lastError
nextAction
needsManualReview

Ovo daje poseban operativni pogled na SEF bez kopanja po generalnom Events tabu.

28. GAS Alerts table

Alerts prima događaje koji zahtevaju pažnju.

Tipični uslovi:

severity = ERROR
severity = CRITICAL
needsManualReview = true
SEF_UNKNOWN
SEF_SEND_EXCEPTION_AFTER_LOCAL_SENDING
SEF_RECOVERY_INVOICE_FAIL
SEF_STATUS_REFRESH_EXCEPTION
BACKUP_FAIL
BANKA_AUTOMAP_ALL_FAIL

Glavna svrha:

kratka lista stvari koje operator/admin mora da pogleda
29. GAS AuditCritical table

AuditCritical prima najosetljivije kritične događaje.

Tipični događaji:

SEF_UNKNOWN
SEF_SEND_EXCEPTION_AFTER_LOCAL_SENDING
SEF_RECOVERY_INVOICE_FAIL
SEF_STARTUP_RECOVERY_FAIL
BANKA_AUTOMAP_ALL_FAIL
critical backup fail

Svrha:

trajan audit trag za događaje koji mogu imati poslovnu, finansijsku ili poresku posledicu
30. GAS Backups table

Backups prima backup evente:

BACKUP_SUCCESS
BACKUP_FAIL

Šalju se podaci kao:

backupType
status
backupFileId
backupLocation
rowsCount
checksum
durationMs
errorMessage

Backup monitoring je best-effort i ne blokira VBA startup.

31. GAS event ID

GAS generiše eventId za uspešno primljen event.

Primer potvrđenog odgovora:

eventId = 81e19cd0-59fa-49c3-adf5-afff7d605ad7

eventId služi kao jedinstveni identifikator monitoring zapisa.

32. GAS timestamp

GAS vraća timestamp u ISO formatu:

2026-05-05T09:55:31.836Z

To omogućava uniformno vreme u monitoring fajlu, nezavisno od lokalnog Excel/VBA formata datuma.

33. GAS component mapping

Za VBA evente komponenta je:

VBA Client

Za specijalizovane tokove komponenta se može izvesti iz:

source
module
eventType
entityType

Efektivno grupisanje:

VBA Client
SEF
Backup
Banka
Faktura
Novac
Otkup
Dokumenti
Startup
34. GAS severity model

Koriste se severity vrednosti:

INFO
WARN
ERROR
CRITICAL

Praktična upotreba:

INFO      = normalan uspešan signal
WARN      = poslovno važno upozorenje, ali nije pad sistema
ERROR     = operacija nije uspela
CRITICAL  = moguć poslovni/porezni/finansijski rizik ili ručni review
35. GAS needsManualReview

Za događaje koji traže ljudsku proveru šalje se:

needsManualReview = true

Najvažniji primeri:

SEF_SEND_EXCEPTION_AFTER_LOCAL_SENDING
SEF_RECOVERY_INVOICE_FAIL
SEF_UNKNOWN
SEF_STATUS_REJECTED

GAS koristi ovo za alert/audit routing.

36. GAS nextAction

Za operativno vođenje koristi se:

nextAction

Primeri:

WAIT
RETRY
MANUAL_REVIEW
CHECK_SEF_PORTAL

Ovo je posebno bitno za SEF monitoring, jer razlikuje tehnički retry od ručne poreske/portal provere.

37. GAS environment

Svaki event nosi:

environment

Vrednosti:

DEV
PROD

VBA čita environment iz:

tblSEFConfig → MONITORING_ENV

Ako nema vrednosti, fallback je:

DEV

Za production monitoring treba:

MONITORING_ENV = PROD
38. GAS public endpoint deployment

GAS mora biti deployovan kao Web App.

VBA koristi isključivo /exec URL.

Posle izmene GAS koda, potreban je novi deployment/redeploy da VBA pogodi novu verziju.

Efektivno stanje potvrđeno testom:

VBA POST → GAS Web App → success:true
39. GAS kao centralni observability sloj

GAS deo sada ima ulogu centralnog observability sloja za:

VBA runtime
SEF integraciju
fakture
novac
otkup
dokumentni lanac
bankarsko mapiranje
backup
startup
health
alerts
audit critical

VBA ostaje business aplikacija.

GAS/Sheets monitoring ostaje eksterni nadzorni sloj.

40. Dopunjen finalni status

Sa GAS delom, ukupna pokrivenost je:

VBA monitoring client: urađen
GAS monitoring ingest: urađen
Secret validation: urađena
Web App endpoint: urađen
Monitoring Sheets routing: urađen
Events log: urađen
Errors log: urađen
Health update: urađen
SEFStatus log: urađen
Backups log: urađen
Alerts log: urađen
AuditCritical log: urađen
End-to-end VBA → GAS test: potvrđen

Finalno stanje:

Monitoring sistem sada ima VBA client, GAS ingest, Google Sheets storage, health, alerts i audit trail.
