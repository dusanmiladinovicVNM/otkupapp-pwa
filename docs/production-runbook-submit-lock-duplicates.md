Production runbook: PWA submit-lock i double-submit duplicate handling
Status: operativni runbook za incidente “kliknuo sam dva puta”, “duplirao se otkup/zbirna/tretman”, “vidim isti red dva puta”, “loša konekcija napravila duplikat”.
Aplikacija: OtkupApp / AgriX PWA
Domen: UI submit → local IndexedDB record → sync engine → GAS/Google → render merge/dedupe → Excel MasterSync
Glavni kod: src/js/utils/async.js, src/js/features/otkup/otkup-form.js, src/js/features/vozac/zbirna.js, src/js/features/kooperant/agromere.js, src/js/utils/sync-engine.js

1. Kada korisnik kaže problem
Tipični incidenti:


“Kliknuo sam Sačuvaj dva puta i sada imam dva otkupa.”


“Zbirna se duplirala.”


“Tretman je upisan dva puta.”


“Vidim isti red dvaput u listi.”


“Bio sam offline/online i posle sync-a pojavila su se dva reda.”


“Dugme je ostalo zaključano.”


“Pisalo je da je čuvanje već u toku.”


“Jedan red je pending, drugi synced.”


“U Google Sheet-u je jedan red, a u PWA vidim dva.”


“U PWA je jedan red, ali u Excelu su dva.”


Prvo pravilo:

Ne zaključuj da je poslovni duplikat dok ne uporediš clientRecordID. Isti clientRecordID znači isti logički zapis. Različiti clientRecordID znači moguća dva stvarna unosa.

Minimalni podaci koje operator mora da prikupi:
Uloga: Otkupac / Vozac / KooperantEntityID:Feature: otkup / zbirna / tretmanVreme unosa:Da li je korisnik kliknuo više puta:Da li je uređaj bio online/offline:Da li je prikazana poruka “već u toku”:Lokalni clientRecordID-jevi:ServerRecordID-jevi:syncStatus za svaki red:Da li red postoji u Google Sheet-u:Da li red postoji u Excel masteru:Da li su poslovna polja identična:

2. Source of truth: gde se gleda
2.1. Prvo mesto: lokalni IndexedDB
Za duplikate u PWA prvo proveri lokalni store:
FeatureStoreGlavni IDOtkupCONFIG.STORE_NAMEclientRecordIDZbirnazbirneclientRecordIDTretmantretmaniclientRecordID
Za svaki sumnjivi red proveri:
clientRecordIDserverRecordIDsyncStatussyncAttemptssyncAttemptAtlastServerStatuslastSyncErrorcreatedAtClientupdatedAtClientsyncedAtupdatedAtServerdeleted
2.2. Drugo mesto: UI runtime lock state
withSubmitLock čuva active lock u:
window.appRuntime.submitLocks
Tipični lock ključevi:
otkup:savezbirna:confirmtretman:save
Ako lock postoji predugo, proveriti:
startedAtreason
2.3. Treće mesto: Google/GAS layer
Ako lokalno postoje sumnjivi redovi, proveri server transport:


Otkup: OTK-* sheet po ClientRecordID;


Zbirna: VOZ-* sheet po ClientRecordID;


Tretman: odgovarajući GAS treatment sheet po ClientRecordID.


2.4. Četvrto mesto: Excel MasterSync
Ako je duplikat u Excelu, proveri:
ClientRecordID / ServerRecordID ako postojibusiness key: datum, entity, partner, kg, artikal, parcelaSyncStatus u Google Sheet-uMasterSync writeback
Ako je Google jedan red, a Excel dva, problem je MasterSync/idempotency, ne UI double-submit.

3. Koji ID pratiš
Primarni ID:
clientRecordID
Sekundarni ID-jevi:
serverRecordIDBrojZbirne ako je zbirnaOtkupID / ZbirnaID / Tretman master ID ako je ušao u Excel/GAS master
Business duplicate poređenje:
Otkup
otkupacIDdatumkooperantIDvrstaVocasortaVocaklasakolicinacenatipAmbalazekolAmbalazevozacIDparcelaIDcreatedAtClient
Zbirna
vozacIDdatumkupacIDkolicinaKlIkolicinaKlIIkolAmbalazeotkupRecordIDscreatedAtClient
Tretman
kooperantIDparcelaIDdatummeraartikalIDkolicinaUpotrebljenavremePocetkavremeZavrsetkageoStart/geoEnd ako postojicreatedAtClient
Incident ticket minimum:
Feature:Store:EntityID:clientRecordID A:clientRecordID B:Same clientRecordID: Da/NesyncStatus A/B:serverRecordID A/B:Google rows:Excel rows:Business fields equal: Da/NeDecision:

4. Kako submit lock treba da radi
4.1. withSubmitLock
withSubmitLock(lockKey, fn, options) radi:


uzima window.appRuntime.submitLocks;


ako lock već postoji za key:


prikaže alreadyMessage ako je definisan;


vrati null;


ne poziva business funkciju;




ako lock ne postoji:


upisuje lock sa startedAt i reason;


disable-uje elemente po data-action selector-u;


dodaje aria-busy=true i aria-disabled=true;




izvršava business funkciju;


u finally briše lock i vraća UI elemente u prethodno stanje.


Ključno:

Lock se briše u finally, pa mora da se očisti i kada business funkcija pukne.

4.2. Otkup save
Public wrapper:
saveOtkup()
Lock:
lockKey = otkup:saveaction = save-otkupalreadyMessage = Čuvanje otkupa je već u toku
Business funkcija:
saveOtkupUnlocked()
Normalan tok:


validira formu;


pravi novi clientRecordID;


upisuje u IndexedDB store CONFIG.STORE_NAME;


status syncStatus = pending;


resetuje formu;


radi safeRefreshAfterSave;


ako je online, pokreće syncQueueSafe('post-save').


4.3. Zbirna confirm
Public wrapper:
confirmZbirna()
Lock:
lockKey = zbirna:confirmaction = confirm-zbirnaalreadyMessage = Kreiranje zbirne je već u toku
Business funkcija:
confirmZbirnaUnlocked()
Normalan tok:


proverava kupca;


uzima današnje neiskorišćene otkupe;


računa kolicinaKlI, kolicinaKlII, kolAmbalaze;


pravi novi clientRecordID;


upisuje u IndexedDB store zbirne;


status syncStatus = pending;


pokreće syncQueueSafe('post-save') ako je online.


4.4. Tretman save
Public wrapper:
agroSaveTretman()
Očekivani lock:
lockKey = tretman:saveaction = agro-save-tretman ili odgovarajući data-action
Business funkcija:
agroSaveTretmanUnlocked()
Normalan tok:


validira parcelu, meru, artikal/količinu ako treba;


proverava meteo/karencu ako je relevantno;


pravi clientRecordID;


upisuje u IndexedDB store tretmani;


status syncStatus = pending;


pokreće sync.



5. Vrste “duplikata”
5.1. UI/render duplikat
Signal:
isti clientRecordID prikazan dva putaGoogle ima jedan redIndexedDB ima jedan red ili local+server merge pokazuje dva alias-a
Zaključak:


ovo nije poslovni duplikat;


problem je render merge/dedupe.


Akcija:


ne brisati poslovni podatak;


proveriti dedupeRecordsForRender i merge alias logiku;


osvežiti view / reload PWA.


5.2. Sync alias duplikat
Signal:
local pending/synced copy i server copy postoje u merge-uisti clientRecordID ili serverRecordID alias
Zaključak:


isti logical record ima dve reprezentacije.


Akcija:


render treba da prikaže jednu;


sync treba da konvergira u synced.


5.3. Local double-submit duplikat
Signal:
dva različita clientRecordID-jacreatedAtClient vrlo blizubusiness fields identičnioba su pending/syncing/synced
Zaključak:


korisnik ili UI je napravio dva lokalna poslovna zapisa.


Akcija:


poslovni owner odlučuje koji ostaje;


tehnički owner proverava zašto lock nije sprečio drugi submit.


5.4. Server duplicate
Signal:
jedan local recordserver/GAS vraća duplicate/existingGoogle ima postojeći red za isti clientRecordID ili poslovni key
Zaključak:


backend idempotency/dedupe radi ili je našao konflikt.


Akcija:


tretirati kao success samo ako je isti clientRecordID ili dokazani isti poslovni zapis;


ako nije, eskalirati.


5.5. Master duplicate
Signal:
Google ima jedan redExcel master ima dva reda
Zaključak:


problem nije submit lock, nego MasterSync idempotency/writeback.


Akcija:


preći na PWA MasterSync runbook.



6. Standardni incident flow
Korak 1: Utvrdi gde se duplikat vidi
Postavi pitanje:
Duplikat je u PWA UI, IndexedDB, Google Sheet-u ili Excel masteru?
Gde se vidiVerovatni domensamo PWA UIrender/merge/dedupeIndexedDB dva redalocal double-submitGoogle dva redasync/server idempotencyExcel dva redaMasterSync duplicatePWA jedan, Excel dvaMasterSyncPWA dva, Google jedanrender/local merge
Korak 2: Uporedi clientRecordID
Za sumnjive redove napravi tabelu:
Row | clientRecordID | serverRecordID | syncStatus | createdAtClient | business fields hashA   | ...            | ...            | ...        | ...             | ...B   | ...            | ...            | ...        | ...             | ...
Tumačenje:
RezultatZnačenjeisti clientRecordIDisti logički zapis, nije poslovni duplikatrazličit clientRecordID, ista poslovna poljamogući double-submitrazličit clientRecordID, različita poljaverovatno dva stvarna unosajedan ima serverRecordID, drugi nemalocal/server merge ili partial sync
Korak 3: Proveri lock state
U DevTools console:
window.appRuntime && window.appRuntime.submitLocks
Ako lock postoji:
lockKeystartedAtreason
Ako lock ne postoji, ali duplikat je nastao, proveri:


da li wrapper za taj flow koristi withSubmitLock;


da li button ima isti data-action koji lock selector očekuje;


da li je business funkcija dostupna direktno iz HTML-a bez wrapper-a;


da li su dva taba otvorena.


Korak 4: Proveri local records
U IndexedDB proveri store:
await dbGetAll(db, CONFIG.STORE_NAME)await dbGetAll(db, 'zbirne')await dbGetAll(db, 'tretmani')
Filter po vremenu i poslovnim poljima.
Korak 5: Proveri server/master
Traži svaki clientRecordID u:
OTK-* sheetVOZ-* sheetTretmani sheetExcel master ako je importovan
Korak 6: Izaberi akciju
SignalAkcijaisti clientRecordID, prikaz dva putarender/dedupe fix; ne dirati podatkedva local pending reda, jedan treba odbacitiposlovni owner odlučuje; tehnički owner markira/delete local ako nije syncovanjedan pending, jedan synced, isti business eventproveriti server; ne syncovati oba dok owner ne odlučioba synced u Googleposlovna korekcija/storno/ignore jednog redaExcel duplicateMasterSync runbooklock zaglavljenreload app; ako se ponavlja, tehnički owner debug finally path

7. Dozvoljene i zabranjene akcije
Dozvoljeno operatoru


zabeležiti sumnjive redove;


reći korisniku da ne unosi ponovo;


proveriti da li je poruka “već u toku” prikazana;


osvežiti ekran ako je UI/render duplikat;


eskalirati sa clientRecordID parovima.


Zabranjeno operatoru


brisati local IndexedDB red naslepo;


ručno menjati clientRecordID;


syncovati oba pending duplikata dok se ne odluči koji je važeći;


ručno brisati Google/Excel red;


tretirati isti clientRecordID kao dva poslovna događaja.


Tehnički owner sme


ručno označiti local duplicate kao deleted ako nije syncovan i poslovni owner odobri;


ručno stopirati sync za jedan pending duplicate;


popraviti direct event binding koji zaobilazi wrapper;


popraviti missing data-action selector;


pojačati server-side idempotency.


Poslovni owner odlučuje


koji od dva različita clientRecordID ostaje;


da li su dva slična unosa stvarno dva događaja;


da li se duplikat stornira, ignoriše ili koriguje;


šta raditi ako je duplikat već u Excel dokumentnom/finansijskom toku.



8. Recovery scenariji
8.1. Korisnik vidi isti red dva puta, ali clientRecordID je isti
Postupak:


Ne brisati ništa.


Proveriti da li je jedan red local copy, drugi server copy.


Reload view / PWA.


Proveriti dedupeRecordsForRender za taj render path.


Ako se ponavlja, tehnički owner popravlja merge aliases.


8.2. Dva različita clientRecordID, oba pending
Postupak:


Uporediti poslovna polja.


Pitati korisnika da li je zaista uneo dva događaja.


Ako je duplikat, ne syncovati oba.


Tehnički owner može local duplicate označiti kao deleted ili ukloniti samo uz export/ticket.


Syncovati samo važeći red.


8.3. Dva različita clientRecordID, jedan pending, jedan synced
Postupak:


Proveriti Google red za synced ID.


Ako synced red je važeći, pending duplikat se ne syncuje.


Ako pending je ispravan, a synced pogrešan, poslovni owner odlučuje korekciju server/master reda.


Ne menjati local status bez server provere.


8.4. Oba duplikata su već synced
Postupak:


Proveriti oba Google redova.


Proveriti da li su ušli u Excel master.


Ako nisu u Excelu, rešiti na Google/MasterSync nivou pre import-a.


Ako jesu u Excelu, domain owner odlučuje storno/korekciju:


Otkup → dokumentni/otkup runbook;


Zbirna → dokumentni chain runbook;


Tretman → agro/knjiga polja korekcija.




8.5. Dugme ostalo zaključano
Postupak:


Proveriti window.appRuntime.submitLocks.


Ako lock postoji, zabeležiti startedAt i reason.


Ako request i save nisu aktivni, reload PWA obično čisti runtime lock.


Proveriti da li je business funkcija završila ali UI nije vraćen.


Ako se ponavlja, tehnički owner proverava da li postoji sync/render code koji blokira main thread ili baca pre wrapper-a.


8.6. Otkup se duplirao
Postupak:


Proveriti CONFIG.STORE_NAME redove.


Uporediti clientRecordID.


Ako su različiti, uporediti: kooperant, datum, vrsta/sorta, klasa, količina, cena, ambalaža, vozač.


Proveriti da li saveOtkup wrapper koristi withSubmitLock i da li HTML button poziva wrapper, ne saveOtkupUnlocked.


Ako je jedan red pending, ne syncovati dok se ne odluči.


Ako su oba u masteru, poslovna korekcija.


8.7. Zbirna se duplirala
Postupak:


Proveriti store zbirne.


Uporediti clientRecordID i otkupRecordIDs.


Ako su otkupRecordIDs isti, to je jak signal duplikata.


Proveriti da li confirmZbirna koristi lock i da li button ne zaobilazi wrapper.


Ako jedna zbirna još nije syncovana, blokirati njen sync.


Ako su obe syncovane, proveriti VOZ-* sheet i MasterSync pre import-a.


Ako su obe importovane, dokumentni chain runbook.


8.8. Tretman se duplirao
Postupak:


Proveriti store tretmani.


Uporediti clientRecordID.


Uporediti parcelu, datum, meru, artikal, količinu, vreme početka/kraja.


Ako je tretman stvarno duplikat i nije syncovan, tehnički owner može local cleanup uz odobrenje.


Ako je syncovan, korekcija u agro/tretman sheet-u uz poslovnu odluku.


8.9. Dva taba PWA otvorena
withSubmitLock je runtime lock u okviru jednog window/app runtime-a. Dva taba mogu imati odvojene runtime lock-ove.
Postupak:


Pitati korisnika da li ima više otvorenih tabova/prozora.


Proveriti deviceID i createdAtClient.


Ako su dva taba napravila dva različita clientRecordID, tretirati kao poslovni duplikat.


Preporuka: koristiti jedan otvoren tab; tehnički owner razmatra cross-tab lock kroz BroadcastChannel, localStorage ili IndexedDB lock.


8.10. Offline pa online duplikat
Postupak:


Proveriti da li je korisnik napravio drugi unos jer prvi nije video kao synced.


Uporediti createdAtClient.


Proveriti sync status oba reda.


Ako je prvi bio pending, drugi nastao ručno, poslovni owner odlučuje.


Ne tretirati sync retry istog clientRecordID kao duplikat.



9. Admin / DevTools komande
9.1. Provera active locks
window.appRuntime && window.appRuntime.submitLocks
9.2. Provera otkup duplikata
const all = await dbGetAll(db, CONFIG.STORE_NAME)all.map(r => ({  id: r.clientRecordID,  status: r.syncStatus,  server: r.serverRecordID,  at: r.createdAtClient,  key: [    r.datum,    r.kooperantID,    r.vrstaVoca,    r.sortaVoca,    r.klasa,    r.kolicina,    r.cena,    r.kolAmbalaze  ].join('|')}))
9.3. Provera zbirna duplikata
const all = await dbGetAll(db, 'zbirne')all.map(r => ({  id: r.clientRecordID,  status: r.syncStatus,  server: r.serverRecordID,  broj: r.brojZbirne,  at: r.createdAtClient,  key: [    r.datum,    r.kupacID,    r.kolicinaKlI,    r.kolicinaKlII,    r.kolAmbalaze,    r.otkupRecordIDs  ].join('|')}))
9.4. Provera tretman duplikata
const all = await dbGetAll(db, 'tretmani')all.map(r => ({  id: r.clientRecordID,  status: r.syncStatus,  server: r.serverRecordID,  at: r.createdAtClient,  key: [    r.datum,    r.parcelaID,    r.mera,    r.artikalID,    r.kolicinaUpotrebljena,    r.vremePocetka,    r.vremeZavrsetka  ].join('|')}))
9.5. Export pre cleanup-a
JSON.stringify(await dbGetAll(db, CONFIG.STORE_NAME), null, 2)JSON.stringify(await dbGetAll(db, 'zbirne'), null, 2)JSON.stringify(await dbGetAll(db, 'tretmani'), null, 2)
9.6. Ručno markiranje local duplicate-a
Samo tehnički owner, samo ako red nije syncovan i poslovni owner odobri:
const r = await dbGet(db, 'zbirne', '<duplicate-clientRecordID>')r.deleted = truer.syncStatus = 'deleted'r.lastServerStatus = 'manual-local-duplicate'r.lastSyncError = 'Marked duplicate after incident review'await dbPut(db, 'zbirne', r)
Ne koristiti ako je red već syncovan na server.

10. Kako sprečavaš duplikate
Postojeće zaštite:


withSubmitLock blokira ponovni klik u istom runtime-u.


Submit dugme se disable-uje i dobija aria-busy / aria-disabled.


Lock se briše u finally.


saveOtkup, confirmZbirna, agroSaveTretman treba da budu public wrappers.


Business funkcije *Unlocked ne treba zvati iz HTML-a direktno.


clientRecordID je local primary key.


Sync engine retry koristi isti clientRecordID.


GAS/Google treba da upisuje/update-uje po clientRecordID.


Render merge koristi dedupe da ne prikaže local/server alias duplo.


Ograničenja:


Lock je client-side runtime lock, nije cross-tab lock.


Lock ne sprečava korisnika da ručno napravi novi unos posle pending statusa.


Lock ne zamenjuje server-side idempotency.


Lock ne popravlja MasterSync duplicate import.


Operativno pravilo:

Retry istog clientRecordID je normalan. Dva različita clientRecordID za isti poslovni događaj su incident dok poslovni owner ne odluči drugačije.


11. Ko donosi odluku
Operator sme sam


prikupiti clientRecordID parove;


proveriti da li je duplikat samo u prikazu;


tražiti od korisnika da ne klikće/unesi ponovo;


reloadovati PWA ako je lock zaglavljen;


eskalirati sa exportom lokalnih redova.


Tehnički owner odlučuje


lokalno brisanje/markiranje unsynced duplikata;


recovery zaglavljenog lock-a;


promenu event binding-a;


dodavanje lock-a na flow koji ga nema;


server-side idempotency fix;


render dedupe fix;


cross-tab locking rešenje.


Poslovni owner odlučuje


da li dva različita clientRecordID predstavljaju isti poslovni događaj;


koji od dva zapisa ostaje;


da li se jedan stornira, ignoriše ili koriguje;


šta sa downstream dokumentima ako je duplikat već importovan.


Niko ne sme bez odobrenja


brisati syncovan server red;


menjati clientRecordID;


ručno setovati synced da bi sakrio duplikat;


importovati oba duplikata u Excel ako je jasno da je isti događaj;


zvati *Unlocked funkcije direktno iz UI-a.



12. Checklist za zatvaranje incidenta
[ ] Identifikovana uloga i feature[ ] Identifikovan store[ ] Izvučeni svi sumnjivi clientRecordID-jevi[ ] Utvrđeno da li su ID-jevi isti ili različiti[ ] Provereni syncStatus/serverRecordID za svaki red[ ] Proveren Google/GAS trag[ ] Proveren Excel master trag ako postoji[ ] Provereno da li je duplikat samo UI/render[ ] Provereno da li submit wrapper koristi withSubmitLock[ ] Provereno da HTML/button ne zove *Unlocked direktno[ ] Ako je business duplikat, postoji odluka poslovnog owner-a[ ] Ako je local cleanup, postoji export i ticket[ ] Ako je server/master cleanup, prebačeno na domain runbook[ ] Korisnik obavešten

13. Primeri odluke
Primer A: Isti clientRecordID se vidi dva puta u PWA
Zaključak: render/merge duplikat, ne poslovni duplikat.
Akcija: ne brisati podatke; reload; tehnički owner proverava dedupe.
Primer B: Dva otkupa imaju različit clientRecordID, isti kooperant/kg/cena, nastali u istoj sekundi
Zaključak: verovatni double-submit.
Akcija: ne syncovati oba ako su pending; poslovni owner bira važeći; tehnički owner proverava lock/event binding.
Primer C: Zbirna se duplirala sa istim otkupRecordIDs
Zaključak: jak signal poslovnog duplikata.
Akcija: zaustaviti sync/import jednog reda; ako je već importovano, dokumentni chain runbook.
Primer D: Jedan tretman pending, drugi synced, ista parcela/artikal/vreme
Zaključak: moguće da je korisnik ponovio unos jer prvi nije video kao synced.
Akcija: proveriti server; ne syncovati pending duplikat dok owner ne odluči.
Primer E: Dugme ostalo disabled posle greške
Zaključak: runtime lock/UI cleanup problem ili long-running function.
Akcija: reload PWA; proveriti submitLocks; tehnički owner debug finally path.

14. Poznate production rupe koje treba zatvoriti


Dodati cross-tab submit lock kroz BroadcastChannel ili IndexedDB lock.


Dodati duplicate warning pre save-a: isti business key u poslednjih X minuta.


Dodati user-visible “Čuvanje je u toku” spinner sa lock key-om.


Dodati centralni SubmitLockEventLog u client ErrorLog za already-running pokušaje.


Dodati server-side idempotency po clientRecordID na svim write endpoint-ima.


Dodati business-key duplicate detection za Otkup/Zbirna/Tretman.


Dodati admin UI za local duplicate cleanup umesto DevTools ručnog rada.


Dodati smoke test za double-tap za Otkupac, Vozač i Kooperant pre svake release sezone.


Dodati test da HTML ne referencira *Unlocked funkcije direktno.


Dodati alert ako se isti business key pojavi sa dva clientRecordID u kratkom roku.


Dodati “pending exists” warning ako korisnik pokušava da unese isti događaj ponovo dok je prvi pending.


Dodati MasterSync blokadu za dva Google reda sa istim business key-em dok owner ne odluči.


Do tada važi konzervativno pravilo:

Submit-lock sprečava slučajni drugi klik, ali production idempotency je clientRecordID + server dedupe + MasterSync dedupe + poslovna odluka za različite ID-jeve.
