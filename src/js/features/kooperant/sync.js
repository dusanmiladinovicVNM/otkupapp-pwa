async function agroSaveTretman() {
    const art = agroState.artikalData;
    const timer = agroState.timerResult;
    const now = new Date();
    const nowIso = now.toISOString();

    let datumBerbeDozvoljeno = '';
    if (agroState.karencaDana > 0) {
        const d = new Date();
        d.setDate(d.getDate() + agroState.karencaDana);
        datumBerbeDozvoljeno = d.toISOString().split('T')[0];
    }

    const record = {
        clientRecordID: (window.crypto && typeof window.crypto.randomUUID === 'function')
            ? window.crypto.randomUUID()
            : ('agro-' + Date.now() + '-' + Math.floor(Math.random() * 1000000)),
        serverRecordID: '',
        createdAtClient: nowIso,
        updatedAtClient: nowIso,
        updatedAtServer: '',
        syncedAt: '',

        kooperantID: CONFIG.ENTITY_ID,
        parcelaID: agroState.parcelaID,
        datum: nowIso.split('T')[0],
        mera: agroState.mera,

        artikalID: art ? art.artikalID : '',
        artikalNaziv: art ? art.naziv : '',
        kolicinaUpotrebljena: agroState.kolicina || '',
        jedinicaMere: art ? art.jm : '',
        dozaPreporucena: agroState.dozaPreporucena || '',
        dozaPrimenjena: agroState.kolicina || '',

        opremaTraktor: agroState.opremaTraktor,
        opremaPrskalica: agroState.opremaPrskalica,
        opremaOstalo: agroState.opremaOstalo,

        karencaDana: agroState.karencaDana || '',
        datumBerbeDozvoljeno: datumBerbeDozvoljeno,

        vremePocetka: timer ? timer.pocetakISO : '',
        vremeZavrsetka: timer ? timer.zavrsetakISO : '',
        trajanjeMinuta: timer ? timer.trajanjeMinuta : '',

        geoLatStart: agroState.geoStart ? agroState.geoStart.lat : '',
        geoLngStart: agroState.geoStart ? agroState.geoStart.lng : '',
        geoLatEnd: agroState.geoEnd ? agroState.geoEnd.lat : '',
        geoLngEnd: agroState.geoEnd ? agroState.geoEnd.lng : '',
        geoAutoDetect: agroState.geoAutoDetect ? 'Da' : '',

        meteoTemp: agroState.meteoSnapshot ? agroState.meteoSnapshot.temp : '',
        meteoWind: agroState.meteoSnapshot ? agroState.meteoSnapshot.wind : '',
        meteoHumidity: agroState.meteoSnapshot ? agroState.meteoSnapshot.humidity : '',
        meteoOverride: agroState.meteoOverride ? 'Da' : '',

        napomena: agroState.napomena,

        syncStatus: 'pending',
        syncAttempts: 0,
        syncAttemptAt: '',
        lastSyncError: '',
        lastServerStatus: '',
        deleted: false,
        entityType: 'tretman',
        schemaVersion: 1
    };

    await dbPut(db, CONFIG.AGRO_STORE, record);
    showToast('Tretman sačuvan!', 'success');

    try {
        await agroLoadIstorija();
    } catch (e) {
        console.error('agroLoadIstorija after save failed:', e);
    }

    if (navigator.onLine) {
        if (typeof syncAgromere === 'function') {
            await syncAgromere();
        }
    }

    agroResetState();
    agroPopulateParcele();
    agroLoadIstorija();
    document.getElementById('agroStep1').style.display = 'block';
    document.getElementById('agroStep2').style.display = 'none';
}
