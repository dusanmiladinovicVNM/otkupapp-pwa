Stvarne otvorene stavke koje sam video
1. confirmZbirnaUnlocked ne filtrira consumedIds u agregaciji — moguć double-count
Pažljivo pogledaj redosled:
javascriptconst todayOtkupi = (vozacOtkupi || []).filter(r =>
    r.datum === today && !consumedIds.has(r.clientRecordID)
);

if (todayOtkupi.length === 0) {
    showToast('Nema otkupa za danas', 'error');
    return;
}

let totalKgI = 0;
let totalKgII = 0;
// ... agregacija nad todayOtkupi ...
Ovo je u redu. Ali, vozacOtkupi je već filtriran u loadVozacData():
javascriptconst consumedIds = getConsumedOtkupIdsFromZbirne(zbirne);
vozacOtkupi = vozacOtkupi.filter(r => !consumedIds.has(r.clientRecordID));
Znači consumedIds se računa dvaput u životnom ciklusu: jednom pri load, drugi put pri confirm. Problem nastaje ako se između load-a i confirm-a desi nešto što promeni _lastMergedZbirne ali ne i vozacOtkupi. Konkretno: drugi vozač završi zbirnu, sync, refresh _lastMergedZbirne ali ne vozacOtkupi. Tada consumedIds u confirmZbirna može biti veći od onoga koji je primenjen na vozacOtkupi — što je defenzivno. Ali obrnuto je opasno: ako _lastMergedZbirne zastari (cache), consumedIds može biti manji nego stvarni — i ti otkupi ulaze u dve zbirne.
Race scenario: vozač pravi zbirnu A, sync se desi, server dodeli, ali _lastMergedZbirne nije refreshed pre nego što vozač pravi zbirnu B sa istim otkupima. Otkupi se referenciraju u obe zbirne. Server-side guard ne postoji — processZbirnaRecord ne proverava da li otkupRecordIDs overlap-uje sa drugim već-postojećim zbirna redovima. Ovo je tihi data corruption rizik.
2. mergeZbirneRecords ne kvalifikuje po vozacID
javascriptfunction mergeZbirneRecords(local, server) {
    return mergeOfflineRecords(local, server, normalizeLocalZbirnaRecord);
}
Pretpostavljam da mergeOfflineRecords matchuje po clientRecordID. Ali server vraća samo zbirne tog vozača (pošto VOZ sheet je per-vozač). Lokalni IDB ima samo zbirne tog vozača (jer je per-uređaj). U trenutnoj realnosti — sve OK.
Ali: kada bi ikada postojao scenario "vozač A loguje se na uređaj koji je ranije koristio vozač B" — IDB i dalje ima zbirne vozača B. getMergedZbirneForVozac bi spojio te tuđe zbirne sa novim server response-om, prikazao u UI-u, a najgore — brojao bi ih u sequence count-u za današnji dan ako predloženi PWA-first BrojZbirne ode u produkciju. Vozač B-jev 7/060526-1 postaje deo seed-a za vozač C-jev sequence.
Da li ovo treba da te brine danas? Ne. Ali "session change clears IDB" pravilo treba dodati pre nego što smo i blizu PWA-first numeracije, i pre prvog hand-off-a uređaja. To je P1 launch issue koji nigde nisam video u changelogu.
3. IsDuplicateZbirnaInMaster linearno skenira ceo tblZbirna po svakom redu
vbFor i = 1 To UBound(data, 1)
    If CStr(Nz(data(i, colCRID), "")) = clientRecordID Then
        IsDuplicateZbirnaInMaster = True
        Exit Function
    End If
Next i
Pozvana je iz ImportOneVOZSheet u petlji nad VOZ redovima. Ako ima 10 vozača × 30 zbirni × 100 dana sezone = 30,000 redova u tblZbirna posle pune sezone, a u jednom sync-u ima 5 novih redova → 150,000 string upoređivanja. Sezonski sustav, ali skalira loše. Sezonski merenje: koliko zbirni će biti u tblZbirna na kraju 1. sezone? Ako je odgovor < 5,000 — non-issue. Ako je 50,000+ — već se oseti.
Lakša varijanta: u ImportOneVOZSheet izgradi Dictionary ClientRecordID → True jednom, pa proveri u O(1). Trivijalna izmena.
4. ImportOneVOZSheet EH: blok ignoriše statusUpdates kolekciju
vbEH:
    LogErr "ImportOneVOZSheet", "Sheet: " & sheetName
    outErrors = outErrors + 1
End Sub
Ako padne usred petlje, lokalna statusUpdates kolekcija sadrži delimične update-ove koji nisu propušteni do WriteBackVOZSyncStatus. To znači: VBA strana je možda već importovala 3 reda u tblZbirna, ali ni jedan red u VOZ sheet-u nema status Synced>Master jer writeback nikad nije pokrenut. Sledeći import će pokušati ponovo, IsDuplicateZbirnaInMaster će vratiti True za ta 3, biće preskočeni, ali VOZ sheet i dalje pokazuje Synced (pending sa PWA strane), nikad Synced>Master (potvrda od master-a) — beskonačno se hvataju kao "skipped" pri svakom ciklusu. Subtilno ali stvarno.
Fix: EH: blok mora da pokuša WriteBackVOZSyncStatus sa onim što je u kolekciji do trenutka greške, ili tx.RollbackTx u parent _TX mora da uključi i Sheet1!F status revert (što ne može jer Sheets su izvan TX scope — ovo je već zabeleženo kao KI-v5.6-01 u tvom AR-u).
5. PWA vozacID u record payload-u nije provereno protiv CONFIG.ENTITY_ID pre slanja
javascriptconst record = {
    // ...
    vozacID: CONFIG.ENTITY_ID,
    // ...
};
GAS proverava:
javascriptif (recordVozacID && recordVozacID !== canonicalVozacID) {
    const err = new Error('VozacID mismatch...');
    // ...
}
I PWA syncStore postavlja entityIdField: 'vozacID' u data.vozacID od trenutnog session-a. Ali — PWA ne validira lokalno da record.vozacID matchuje CONFIG.ENTITY_ID pre dbPut-a. Ako se CONFIG.ENTITY_ID promeni između confirmZbirna poziva i sledećeg sync-a (npr. logout i drugi vozač se uloguje na isti uređaj — vidi tačku #2 gore), zbirna ostaje u IDB sa starim vozacID. Sync će zatim ili padne (mismatch) ili — gore — uspe ako entityID ne stigne do GAS-a iz session-a već iz record-a.
Ovo se vezuje za session change → IDB clear pravilo iz tačke #2.
6. _lastMergedZbirne cache nikad ne expire-uje
javascriptlet _lastMergedZbirne = null;
Module-scoped, popunjen u loadVozacData() i loadVozacZbirne(). Korišćen u confirmZbirnaUnlocked kao izvor za consumedIds (tačka #1). Ali ako vozač otvori app, snimi 1 zbirnu (sync OK), zatim ne pozove ni loadVozacData ni loadVozacZbirne (npr. ostane na zbirna create view-u i pravi drugu) — _lastMergedZbirne može biti stale iz pre snimanja prve zbirne. Ovo nije apokaliptično jer dbGetAll u getMergedZbirneForVozac pruža kompenzaciju kad se ipak pozove. Ali pattern "module-scope mutable cache without invalidation hooks" je dug koji se kumulativno akumulira.
7. GAS processZbirnaRecord ne validira da otkupRecordIDs referenciraju otkupe tog vozača
Linija u GAS-u koju vidim:
javascriptOtkupRecordIDs: record.otkupRecordIDs || ''
Upisuje se kao opaque string. Ne postoji provera da svi ID-evi u tom stringu pripadaju otkupima čiji je VozacID jednak canonicalVozacID. Maliciozni ili buggy klijent može poslati otkupRecordIDs koji uključuje otkupe drugih vozača. Sledeći put kad VBA cascade LinkZbirnaToOtkupAndOtpremnica propagira BrojZbirne, postavi BrojZbirne tuđim otkupima.
Manje verovatno u praksi (klijent mora aktivno da pošalje pogrešne ID-eve), ali authorization model je nepotpun na ovom polju. Hard-end fix: GAS lookup za svaki ID, validacija VozacID match-a, odbij ako bilo koji ne match-uje.
8. formatDate(today, 'ddmmyy') poziv koji sam ti predložio — ne postoji u tvom codebase-u
Vraćam se na svoj sopstveni predlog za PWA generisanje BrojZbirne. Predložio sam formatDdmmyy(today) helper koji ne postoji. Trebao sam da kažem: dodaj ga, evo skeleton-a. To sam i učinio kasnije, ali tek u trećem ili četvrtom predlogu.
To je signal da prolazim kroz tvoj kod brže nego što treba. Ako ima drugih helper-a koje sam pretpostavio da postoje a zapravo ne postoje, ne bi me iznenadilo — posebno ako sam u nekom trenutku predložio nešto kao "samo izmeni X" gde X u tvom codebase-u nije baš taj X.
9. SaveZbirnaMulti_TX zove SaveZbirna dva puta — GetNextID se zove dvaput
vbresultI = SaveZbirna(... KLASA_I)
' ...
resultII = SaveZbirna(... KLASA_II)
Svaki SaveZbirna zove GetNextID(TBL_ZBIRNA, COL_ZBR_ID, "ZBR-"). Pretpostavljam da GetNextID čita postojeću tabelu i daje sledeći broj. Između prvog i drugog poziva, tblZbirna ima novi red sa Klasa I — pa drugi poziv vraća N+1. Da li je to željeno ponašanje? Da li ZBR-00042 (Klasa I) i ZBR-00043 (Klasa II) trebaju biti različiti ZbirnaID-evi za logički isti business document?
Iz schema komentara: "one BrojZbirne may span multiple rows by class" (linija 1579 u Reference). OK, znači po dizajnu — iste BrojZbirne ali različiti ZbirnaID-evi. To znači da je ZbirnaID u stvari row-level identifier, a ne document-level identifier. To je iznenađujuće naming convention. BrojZbirne je document-level, ZbirnaID je row-level. Ako negde u kodu postoji "lookup zbirna by ZbirnaID" očekivanje single-row return-a, a u Klasa I+II slučaju ima dva reda — moguć subtle bug. Nisam našao konkretan slučaj, ali pattern me bune.
10. PWA record.entityType = 'zbirna', schemaVersion = 1 — nigde se ne koristi za migration
Vidim ova polja u confirmZbirnaUnlocked:
javascriptdeleted: false,
entityType: 'zbirna',
schemaVersion: 1
Nigde u kodu koji si mi poslao ne vidim schemaVersion proveru. Ovo je infrastructure za buduću migraciju koja nije zapravo aktivna. Nije bag, samo dead-but-helpful field. Ako planiraš migration runner kasnije, dobro je. Ako ne — bezvredan overhead u IDB-u.
11. Memory edits koje sam pročitao na startu — dva su zastarela ako bi v6.15 prošao
Zapazio sam u tvojoj memoriji:

"VBA / ambalaza (packaging) chain — ... Remaining fix: SaveOtkup in modOtkup still needs Stanica Ulaz call added."
"tblZbirna / tblOtkup schema: Both tables require ClientRecordID and SyncSource columns at end. ClientRecordID and SyncSource as columns 15–16 in rowData are not yet in the AR canonical tblZbirna schema (14 columns) — requires schema extension or removal."

Drugu si možda već rešio jer u v6.14 Reference vidim 16-kolonsku kanonu. Prva (Stanica Ulaz call) — ne mogu da potvrdim, ali ako je tako — to je p1 ambalaza-tačnost issue koji je veći nego BrojZbirne diskusija.
