// ============================================================
// MANAGEMENT: KOOPERANTI
// ============================================================
function populateMgmtStanice() {
    const stanice = stammdaten.stanice || [];
    const fallbackIDs = new Set();
    if (stanice.length === 0) {
        (stammdaten.kooperanti || []).forEach(k => { if (k.StanicaID) fallbackIDs.add(k.StanicaID); });
    }
    ['mgmtStanica', 'mgmtOtkupiStanica'].forEach(selId => {
        const sel = document.getElementById(selId);
        if (!sel) return;
        sel.innerHTML = '<option value="">-- Izaberi stanicu --</option>';
        if (stanice.length > 0) {
            stanice.forEach(s => {
                const o = document.createElement('option');
                o.value = s.StanicaID;
                o.textContent = (s.Naziv || s.Mesto || s.StanicaID) + ' (' + s.StanicaID + ')';
                sel.appendChild(o);
            });
        } else {
            fallbackIDs.forEach(id => {
                const o = document.createElement('option');
                o.value = id; o.textContent = id;
                sel.appendChild(o);
            });
        }
    });
}

function onMgmtStanicaChange() {
    const stanicaID = document.getElementById('mgmtStanica').value;
    const sel = document.getElementById('mgmtKooperant');
    sel.innerHTML = '<option value="">-- Izaberi kooperanta --</option>';
    document.getElementById('mgmtKarticaHeader').style.display = 'none';
    document.getElementById('mgmtKarticaList').innerHTML = '';
    if (!stanicaID) return;
    (stammdaten.kooperanti || []).filter(k => k.StanicaID === stanicaID).forEach(k => {
        const o = document.createElement('option'); o.value = k.KooperantID;
        o.textContent = k.Ime + ' ' + k.Prezime + ' (' + k.KooperantID + ')'; sel.appendChild(o);
    });
}

async function onMgmtKooperantChange() {
    const koopID = document.getElementById('mgmtKooperant').value;
    if (!koopID) { document.getElementById('mgmtKarticaHeader').style.display = 'none'; document.getElementById('mgmtKarticaList').innerHTML = ''; return; }
    const koop = (stammdaten.kooperanti || []).find(k => k.KooperantID === koopID);
    document.getElementById('mgmtKarticaName').textContent = koop ? koop.Ime + ' ' + koop.Prezime : koopID;
    document.getElementById('mgmtKarticaID').textContent = koopID;
    document.getElementById('mgmtKarticaHeader').style.display = 'block';

    let records = [];
    if (mgmtData && mgmtData.kartice) {
        records = mgmtData.kartice.filter(r => r.KooperantID === koopID && r.Opis !== 'UKUPNO');
    } else {
        const json = await safeAsync(async () => {
            return await apiFetch('action=getMgmtKartica&kooperantID=' + encodeURIComponent(koopID));
        }, 'Greška pri učitavanju kartice kooperanta');

        if (json && json.success && json.records) {
            records = json.records.filter(r => r.Opis !== 'UKUPNO');
        }
    }
    if (records.length === 0) {
        document.getElementById('mgmtKarticaList').innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:20px;">Nema podataka</p>';
        ['mgmtKarticaZad','mgmtKarticaRaz','mgmtKarticaSaldo'].forEach(id => document.getElementById(id).textContent = '0');
        return;
    }
    let zad = 0, raz = 0;
    document.getElementById('mgmtKarticaList').innerHTML = records.map(r => {
        const z = parseFloat(r.Zaduzenje)||0, ra = parseFloat(r.Razduzenje)||0, s = parseFloat(r.Saldo)||0;
        zad += z; raz += ra;
        return `<div class="queue-item" style="border-left-color:${z>0?'var(--danger)':'var(--success)'};">
            <div class="qi-header"><span class="qi-koop">${escapeHtml(r.BrojDok||'')}</span><span class="qi-time">${escapeHtml(fmtDate(r.Datum))}</span></div>
            <div class="qi-detail">${escapeHtml(r.Opis||'')}</div>
            <div class="qi-detail" style="font-size:12px;margin-top:2px;">
                ${z>0?'<span style="color:var(--danger);">Zaduž: '+z.toLocaleString('sr')+'</span> ':''}
                ${ra>0?'<span style="color:var(--success);">Razduž: '+ra.toLocaleString('sr')+'</span> ':''}
                | Saldo: <strong>${s.toLocaleString('sr')}</strong></div></div>`;
    }).join('');
    document.getElementById('mgmtKarticaZad').textContent = zad.toLocaleString('sr');
    document.getElementById('mgmtKarticaRaz').textContent = raz.toLocaleString('sr');
    document.getElementById('mgmtKarticaSaldo').textContent = (zad - raz).toLocaleString('sr');
}

function loadMgmtKoopSaldo() {
    const kartice = (mgmtData && mgmtData.kartice) ? mgmtData.kartice : [];
    const list = document.getElementById('mgmtKoopSaldoList');
    const totals = kartice.filter(r => r.Opis === 'UKUPNO');
    if (totals.length === 0) { list.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:20px;">Nema podataka</p>'; return; }
    list.innerHTML = totals.map(r => {
        const koop = (stammdaten.kooperanti || []).find(k => k.KooperantID === r.KooperantID);
        const name = koop ? koop.Ime + ' ' + koop.Prezime : r.KooperantID;
        const saldo = parseFloat(r.Saldo)||0;
        const zad = parseFloat(r.Zaduzenje)||0, raz = parseFloat(r.Razduzenje)||0;
        const bc = saldo > 0 ? 'var(--warning)' : 'var(--success)';
        return `<div class="queue-item" style="border-left-color:${bc};">
            <div class="qi-header"><span class="qi-koop">${escapeHtml(name)}</span><span class="qi-time">Saldo: ${saldo.toLocaleString('sr')} RSD</span></div>
            <div class="qi-detail">Zaduž: ${zad.toLocaleString('sr')} | Razduž: ${raz.toLocaleString('sr')}</div></div>`;
    }).join('');
}

function loadMgmtKoopPregled() {
    // Popuni dropdown stanica
    const sel = document.getElementById('mgmtPregledStanica');
    if (sel.options.length <= 1) {
        const stanice = new Set();
        const data = (mgmtData && mgmtData.saldoOMDetail) ? mgmtData.saldoOMDetail : [];
        data.forEach(r => { if (r.StanicaID) stanice.add(r.StanicaID); });
        stanice.forEach(s => { const o = document.createElement('option'); o.value = s; const st = (stammdaten.stanice||[]).find(x => x.StanicaID === s); o.textContent = st ? (st.Naziv||st.Mesto||s)+' ('+s+')' : s; sel.appendChild(o); });
    }
    renderMgmtKoopPregled();
}

function renderMgmtKoopPregled() {
    const stanicaFilter = document.getElementById('mgmtPregledStanica').value;
    let records = (mgmtData && mgmtData.saldoOMDetail) ? mgmtData.saldoOMDetail : [];
    if (stanicaFilter) records = records.filter(r => r.StanicaID === stanicaFilter);
    
    const list = document.getElementById('mgmtKoopPregledList');
    if (records.length === 0) { list.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:20px;">Nema podataka</p>'; return; }
    
    // Sortiraj po stanici pa po imenu
    records.sort((a, b) => (a.StanicaID||'').localeCompare(b.StanicaID||'') || (a.Kooperant||'').localeCompare(b.Kooperant||''));
    
    // Totali
    let totKg = 0, totVr = 0, totIsp = 0, totAgro = 0, totSaldo = 0, totAmb = 0;
    records.forEach(r => {
        totKg += parseFloat(r.Kolicina)||0; totVr += parseFloat(r.Vrednost)||0;
        totIsp += parseFloat(r.Isplaceno)||0; totAgro += parseFloat(r.AgroZaduzenje)||0;
        totSaldo += parseFloat(r.Saldo)||0; totAmb += parseFloat(r.Ambalaza)||0;
    });
    
    list.innerHTML = `
        <div class="stats-grid" style="margin-bottom:12px;">
            <div class="stat-card"><div class="stat-value" style="font-size:20px;">${totKg.toLocaleString('sr')}</div><div class="stat-label">Ukupno kg</div></div>
            <div class="stat-card"><div class="stat-value" style="font-size:20px;">${totVr.toLocaleString('sr')}</div><div class="stat-label">Vrednost</div></div>
            <div class="stat-card"><div class="stat-value" style="font-size:20px;">${totIsp.toLocaleString('sr')}</div><div class="stat-label">Isplaćeno</div></div>
            <div class="stat-card"><div class="stat-value" style="font-size:20px;">${totSaldo.toLocaleString('sr')}</div><div class="stat-label">Saldo</div></div>
        </div>
        ${records.map(r => {
            const kg = parseFloat(r.Kolicina)||0;
            const vr = parseFloat(r.Vrednost)||0;
            const isp = parseFloat(r.Isplaceno)||0;
            const agro = parseFloat(r.AgroZaduzenje)||0;
            const saldo = parseFloat(r.Saldo)||0;
            const amb = parseFloat(r.Ambalaza)||0;
            const bc = saldo > 0 ? 'var(--warning)' : 'var(--success)';
            return `<div class="queue-item" style="border-left-color:${bc};">
                <div class="qi-header"><span class="qi-koop">${escapeHtml(r.Kooperant||r.KooperantID)}</span><span class="qi-time">${escapeHtml(fmtStanica(r.StanicaID))}</span></div>
                <div class="qi-detail">
                    ${kg.toLocaleString('sr')} kg | Vrednost: ${vr.toLocaleString('sr')}
                </div>
                <div class="qi-detail" style="font-size:12px;margin-top:2px;">
                    Isplaćeno: ${isp.toLocaleString('sr')} | Agro: ${agro.toLocaleString('sr')} | Amb: ${amb.toLocaleString('sr')}
                </div>
                <div class="qi-detail" style="font-size:13px;margin-top:4px;font-weight:600;">
                    Saldo: <span style="color:${saldo>0?'var(--danger)':'var(--success)'};">${saldo.toLocaleString('sr')} RSD</span>
                </div>
            </div>`;
        }).join('')}`;
}
