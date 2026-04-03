// ============================================================
// OTKUPNI LIST + SIGNATURE PAD
// ============================================================
function showOtkupniList(record) {
    const config = stammdaten.config || [];
    const gv = k => { const c = config.find(c => c.Parameter === k); return c ? c.Vrednost : ''; };
    const koop = (stammdaten.kooperanti || []).find(k => k.KooperantID === record.kooperantID) || {};
    const vrednostNum = record.kolicina * record.cena;
    const pdvStopa = parseFloat(gv('OtkupPDVStopa')) || 8;
    const pdvIznos = Math.round(vrednostNum * pdvStopa / 100);
    const ukupno = vrednostNum + pdvIznos;

    let modal = document.getElementById('otkupniListModal');
    if (!modal) { modal = document.createElement('div'); modal.id = 'otkupniListModal'; modal.style.cssText = 'position:fixed;top:0;left:0;width:100%;height:100%;background:white;z-index:9999;overflow-y:auto;'; document.body.appendChild(modal); }

    modal.innerHTML = `<div style="padding:16px;max-width:420px;margin:0 auto;font-family:sans-serif;">
        <div style="text-align:center;border-bottom:2px solid #333;padding-bottom:10px;margin-bottom:12px;">
            <div style="font-size:18px;font-weight:700;">${gv('SELLER_NAME')}</div>
            <div style="font-size:12px;color:#666;">${gv('SELLER_STREET')}, ${gv('SELLER_CITY')} ${gv('SELLER_POSTAL_CODE')}</div>
            <div style="font-size:12px;color:#666;">PIB: ${gv('SELLER_PIB')} | MB: ${gv('SELLER_MATICNI_BROJ')}</div>
            <div style="font-size:12px;color:#666;">TR: ${gv('SELLER_ACCOUNT')}</div>
        </div>
        <h2 style="text-align:center;margin-bottom:14px;font-size:20px;">OTKUPNI LIST</h2>
        <div style="background:#f5f5f0;padding:10px;border-radius:8px;margin-bottom:12px;font-size:13px;">
            <div><strong>${koop.Ime || ''} ${koop.Prezime || ''}</strong></div>
            <div>${koop.Adresa || ''}, ${koop.Mesto || ''}</div>
            <div>JMBG: ${koop.JMBG || '________'} | BPG: ${koop.BPGBroj || '________'}</div>
        </div>
        <table style="width:100%;border-collapse:collapse;font-size:14px;">
            <tr><td style="padding:6px;color:#666;width:40%;">Datum:</td><td style="padding:6px;font-weight:600;">${record.datum}</td></tr>
            <tr><td style="padding:6px;color:#666;">Proizvod:</td><td style="padding:6px;">${record.vrstaVoca} ${record.sortaVoca || ''}</td></tr>
            <tr><td style="padding:6px;color:#666;">Klasa:</td><td style="padding:6px;">${record.klasa}</td></tr>
            <tr><td style="padding:6px;color:#666;">Količina:</td><td style="padding:6px;font-weight:600;">${record.kolicina} kg</td></tr>
            <tr><td style="padding:6px;color:#666;">Cena:</td><td style="padding:6px;">${record.cena} RSD/kg</td></tr>
            <tr style="border-top:1px solid #ccc;"><td style="padding:6px;color:#666;">Vrednost:</td><td style="padding:6px;font-weight:600;">${vrednostNum.toLocaleString('sr')} RSD</td></tr>
            ${pdvStopa > 0 ? '<tr><td style="padding:6px;color:#666;">PDV naknada (' + pdvStopa + '%):</td><td style="padding:6px;">' + pdvIznos.toLocaleString('sr') + ' RSD</td></tr>' : ''}
            <tr style="border-top:2px solid #333;"><td style="padding:8px;color:#666;">ZA ISPLATU:</td><td style="padding:8px;font-weight:700;font-size:18px;">${ukupno.toLocaleString('sr')} RSD</td></tr>
            <tr><td style="padding:6px;color:#666;">Ambalaža:</td><td style="padding:6px;">${record.kolAmbalaze} kom</td></tr>
            ${record.parcelaID ? '<tr><td style="padding:6px;color:#666;">Parcela:</td><td style="padding:6px;">' + record.parcelaID + '</td></tr>' : ''}
            <tr><td style="padding:6px;color:#666;">Rok isplate:</td><td style="padding:6px;">${gv('OtkupRokIsplate') || 'Po dogovoru'}</td></tr>
        </table>
        <div style="margin-top:20px;">
            <div style="margin-bottom:16px;"><div style="font-size:12px;color:#666;margin-bottom:4px;">Potpis otkupljivača:</div><canvas id="sigOtkupac" width="720" height="200" style="border:1px solid #ccc;border-radius:6px;width:100%;height:80px;touch-action:none;"></canvas></div>
            <div style="margin-bottom:16px;"><div style="font-size:12px;color:#666;margin-bottom:4px;">Potpis kooperanta:</div><canvas id="sigKooperant" width="720" height="200" style="border:1px solid #ccc;border-radius:6px;width:100%;height:80px;touch-action:none;"></canvas></div>
        </div>
        <div style="text-align:center;margin-top:16px;display:flex;gap:8px;">
            <button onclick="clearSignature('sigOtkupac');clearSignature('sigKooperant')" style="flex:1;padding:12px;font-size:14px;background:#f5f5f0;color:#666;border:1px solid #ccc;border-radius:8px;">Obriši</button>
            <button onclick="saveOtkupniListWithSignatures('${record.clientRecordID}')" style="flex:1;padding:12px;font-size:14px;background:var(--primary);color:white;border:none;border-radius:8px;">Potvrdi</button>
            <button onclick="window.print()" style="flex:1;padding:12px;font-size:14px;background:var(--accent);color:white;border:none;border-radius:8px;">Štampaj</button>
        </div>
        <button onclick="savePdfToDrive('${record.clientRecordID}')" style="width:100%;padding:12px;margin-top:8px;font-size:14px;background:#2196F3;color:white;border:none;border-radius:8px;">📄 Sačuvaj PDF na Drive</button>
        <button onclick="document.getElementById('otkupniListModal').style.display='none'" style="width:100%;padding:10px;margin-top:8px;font-size:14px;background:none;color:#666;border:1px solid #ccc;border-radius:8px;">Zatvori</button>
    </div>`;
    modal.style.display = 'block';
    setTimeout(() => { initSignaturePad('sigOtkupac'); initSignaturePad('sigKooperant'); }, 100);
}

