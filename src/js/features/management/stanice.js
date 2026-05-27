// ============================================================
// MANAGEMENT: STANICE
// ============================================================
async function loadMgmtOtkupi() {
    if (typeof ensureMgmtSection === 'function') {
        await ensureMgmtSection('otkupiAll', 'getMgmtOtkupiAll');
    }

    
    const stanicaID = document.getElementById('mgmtOtkupiStanica').value;
    if (!stanicaID) { document.getElementById('mgmtOtkupiList').innerHTML = ''; return; }
    const od = document.getElementById('mgmtOtkupiOd').value, doo = document.getElementById('mgmtOtkupiDo').value;
    
    let records = [];
    if (mgmtData && mgmtData.otkupiAll) {
        records = mgmtData.otkupiAll
            .filter(r => (r._sheetName === 'OTK-' + stanicaID) || (r.OtkupacID === stanicaID))
            .map(r => ({
                datum: fmtDate(r.Datum), kooperantName: r.KooperantName || r.KooperantID || '',
                vrstaVoca: r.VrstaVoca || '', klasa: r.Klasa || 'I',
                kolicina: parseFloat(r.Kolicina) || 0, cena: parseFloat(r.Cena) || 0
            }));
    } else {
        const json = await safeAsync(async () => {
            return await apiFetch('action=getMgmtOtkupiByStanica&stanicaID=' + encodeURIComponent(stanicaID));
        }, 'Greška pri učitavanju otkupa za stanicu');

        if (json && json.success && json.records) {
            records = json.records.map(r => ({
                datum: fmtDate(r.Datum),
                kooperantName: r.KooperantName || r.KooperantID || '',
                vrstaVoca: r.VrstaVoca || '',
                klasa: r.Klasa || 'I',
                kolicina: parseFloat(r.Kolicina) || 0,
                cena: parseFloat(r.Cena) || 0
            }));
        }
    }
    
    if (od) records = records.filter(r => r.datum >= od);
    if (doo) records = records.filter(r => r.datum <= doo);
    records.sort((a, b) => b.datum.localeCompare(a.datum));
    const kg = records.reduce((s, r) => s + (r.kolicina || 0), 0);
    const vr = records.reduce((s, r) => s + (r.kolicina || 0) * (r.cena || 0), 0);
    document.getElementById('mgmtOtkupiCount').textContent = records.length;
    document.getElementById('mgmtOtkupiKg').textContent = kg.toLocaleString('sr');
    document.getElementById('mgmtOtkupiVrednost').textContent = vr.toLocaleString('sr');
    document.getElementById('mgmtOtkupiKoop').textContent = new Set(records.map(r => r.kooperantName)).size;
    const list = document.getElementById('mgmtOtkupiList');
    if (records.length === 0) {
        list.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:20px;">Nema otkupa</p>';
        return;
    }
    // Render as redesigned table component
    list.innerHTML = `
        <div class="tbl">
            <div class="tbl__head">
                <div>Kooperant</div>
                <div>Vrsta i klasa</div>
                <div style="text-align:right;">Količina</div>
                <div style="text-align:right;">Cena</div>
                <div style="text-align:right;">Vrednost</div>
                <div>Datum</div>
            </div>
            ${records.map(r => {
                const v = ((r.kolicina || 0) * (r.cena || 0));
                return `
                    <div class="tbl__row">
                        <div>
                            <div class="tbl__cell-main">${escapeHtml(r.kooperantName)}</div>
                        </div>
                        <div>
                            <div class="tbl__cell-main">${escapeHtml(r.vrstaVoca)}</div>
                            <div class="tbl__cell-sub">Klasa ${escapeHtml(r.klasa)}</div>
                        </div>
                        <div style="text-align:right;">
                            <div class="tbl__num">${r.kolicina.toLocaleString('sr')}<span class="u">kg</span></div>
                        </div>
                        <div style="text-align:right;">
                            <div class="tbl__num">${r.cena.toLocaleString('sr')}<span class="u">RSD</span></div>
                        </div>
                        <div style="text-align:right;">
                            <div class="tbl__num">${v.toLocaleString('sr')}<span class="u">RSD</span></div>
                        </div>
                        <div>
                            <div class="tbl__cell-sub">${escapeHtml(r.datum)}</div>
                        </div>
                    </div>
                `;
            }).join('')}
        </div>
    `;
}

function loadMgmtSaldoOM() {
    const records = (mgmtData && mgmtData.saldoOM) ? mgmtData.saldoOM : [];
    const list = document.getElementById('mgmtSaldoOMList');
    if (records.length === 0) { list.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:20px;">Nema podataka</p>'; return; }
    list.innerHTML = records.map(r => {
        const saldo = parseFloat(r.Saldo)||0;
        const bc = saldo > 0 ? 'var(--warning)' : 'var(--success)';
        return `<div class="queue-item" style="border-left-color:${bc};">
            <div class="qi-header"><span class="qi-koop">${escapeHtml(fmtStanica(r.StanicaID || r.Stanica))}</span><span class="qi-time">Saldo: ${saldo.toLocaleString('sr')} RSD</span></div>
            <div class="qi-detail">Avans: ${(parseFloat(r.Avans)||0).toLocaleString('sr')} | Isplaceno: ${(parseFloat(r.Isplaceno)||0).toLocaleString('sr')}</div></div>`;
    }).join('');
}

function loadMgmtOtkupPoOM() {
    const records = (mgmtData && mgmtData.otkupPoOM) ? mgmtData.otkupPoOM : [];
    const list = document.getElementById('mgmtOtkupPoOMList');
    if (records.length === 0) { list.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:20px;">Nema podataka</p>'; return; }
    const grouped = {};
    records.forEach(r => {
        const s = r.StanicaID||'?';
        if (!grouped[s]) grouped[s] = { items: [], totalKg: 0, totalAmb: 0, totalVr: 0 };
        const kg = parseFloat(r.Kolicina)||0, amb = parseFloat(r.Ambalaza)||0, vr = parseFloat(r.Vrednost)||0;
        grouped[s].items.push(r); grouped[s].totalKg += kg; grouped[s].totalAmb += amb; grouped[s].totalVr += vr;
    });
    list.innerHTML = Object.entries(grouped).map(([stanica, g]) =>
        `<div style="background:white;border-radius:var(--radius);padding:14px;margin-bottom:12px;box-shadow:0 1px 3px rgba(0,0,0,0.08);border-left:4px solid var(--primary);">
            <div style="display:flex;justify-content:space-between;margin-bottom:8px;"><strong style="color:var(--primary);font-size:16px;">${escapeHtml(fmtStanica(stanica))}</strong><span style="font-size:13px;font-weight:600;">${g.totalKg.toLocaleString('sr')} kg</span></div>
            <div style="font-size:12px;color:var(--text-muted);margin-bottom:8px;">Amb: ${g.totalAmb.toLocaleString('sr')} | Vrednost: ${g.totalVr.toLocaleString('sr')} RSD</div>
            ${g.items.map(r => { const kg=parseFloat(r.Kolicina)||0,amb=parseFloat(r.Ambalaza)||0,vr=parseFloat(r.Vrednost)||0; return `<div style="padding:4px 0;font-size:12px;border-top:1px solid #eee;display:flex;justify-content:space-between;"><span>${escapeHtml(r.VrstaVoca)} ${escapeHtml(r.Klasa)}</span><span>${kg.toLocaleString('sr')} kg | ${amb.toLocaleString('sr')} amb | ${vr.toLocaleString('sr')} RSD | ${escapeHtml(r.BrojOtkupa||0)} otk.</span></div>`; }).join('')}
        </div>`).join('');
}

async function refreshMgmtOtkupiLive() {
    window.mgmtData = window.mgmtData || {};
    delete window.mgmtData.otkupiAll;
    mgmtData = window.mgmtData;

    if (typeof ensureMgmtSection === 'function') {
        await ensureMgmtSection('otkupiAll', 'getMgmtOtkupiAll', 'includeLive=1');
    }

    if (typeof loadMgmtOtkupi === 'function') {
        await loadMgmtOtkupi();
    }
}

async function loadMgmtOtkupUzivo() {
    if (typeof ensureMgmtSection === 'function') {
        await ensureMgmtSection('otkupiAll', 'getMgmtOtkupiAll');
    }

    const today = (typeof mgmtDashTodayISO === 'function') ? mgmtDashTodayISO() : new Date().toISOString().slice(0, 10);
    const od = document.getElementById('mgmtUzivoOd');
    const doo = document.getElementById('mgmtUzivoDo');
    const stanicaSel = document.getElementById('mgmtUzivoStanica');

    // Populate station filter
    if (stanicaSel && stanicaSel.options.length <= 1) {
        const stanice = stammdaten ? (stammdaten.stanice || []) : [];
        stanice.forEach(s => {
            const o = document.createElement('option');
            o.value = s.StanicaID;
            o.textContent = (s.Naziv || s.Mesto || s.StanicaID) + ' (' + s.StanicaID + ')';
            stanicaSel.appendChild(o);
        });
    }

    let records = (window.mgmtData && Array.isArray(window.mgmtData.otkupiAll))
        ? window.mgmtData.otkupiAll
        : [];

    // Date normalizer
    const fmtD = (v) => {
        if (!v) return '';
        if (typeof mgmtDashFmtDate === 'function') return mgmtDashFmtDate(v);
        const s = String(v).trim();
        const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
        return m ? `${m[1]}-${m[2]}-${m[3]}` : '';
    };

    // Apply filters
    const odVal = od ? od.value : today;
    const dooVal = doo ? doo.value : today;
    const stanicaFilter = stanicaSel ? stanicaSel.value : '';

    records = records.filter(r => {
        const d = fmtD(r.Datum);
        if (odVal && d < odVal) return false;
        if (dooVal && d > dooVal) return false;
        if (stanicaFilter) {
            const sid = r.OtkupacID || r.StanicaID || (r._sheetName ? r._sheetName.replace('OTK-', '') : '');
            if (sid !== stanicaFilter) return false;
        }
        return true;
    });

    // Sort by date desc, then by time if available
    records.sort((a, b) => {
        const da = fmtD(a.Datum) || '';
        const db = fmtD(b.Datum) || '';
        return db.localeCompare(da);
    });

    // Compute summaries
    const kg = records.reduce((s, r) => s + (parseFloat(r.Kolicina) || 0), 0);
    const vr = records.reduce((s, r) => s + ((parseFloat(r.Kolicina) || 0) * (parseFloat(r.Cena) || 0)), 0);
    const koops = new Set(records.map(r => r.KooperantID || r.KooperantName || '').filter(Boolean));

    const setTxt = (id, v) => { const el = document.getElementById(id); if (el) el.textContent = v; };
    setTxt('mgmtUzivoCount', records.length.toLocaleString('sr'));
    setTxt('mgmtUzivoKg', kg.toLocaleString('sr'));
    setTxt('mgmtUzivoVrednost', vr.toLocaleString('sr'));
    setTxt('mgmtUzivoKoop', koops.size.toLocaleString('sr'));

    const list = document.getElementById('mgmtUzivoList');
    if (!list) return;

    if (records.length === 0) {
        list.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:20px;">Nema otkupa za izabrani period</p>';
        return;
    }

    // Render as tbl component
    list.innerHTML = `
        <div class="tbl">
            <div class="tbl__head">
                <div>Kooperant</div>
                <div>Vrsta i klasa</div>
                <div>Stanica</div>
                <div style="text-align:right;">Količina</div>
                <div style="text-align:right;">Cena</div>
                <div>Datum</div>
            </div>
            ${records.map(r => {
                const kg = parseFloat(r.Kolicina) || 0;
                const cena = parseFloat(r.Cena) || 0;
                const sid = r.OtkupacID || r.StanicaID || (r._sheetName ? r._sheetName.replace('OTK-', '') : '');
                const stanicaName = typeof fmtStanica === 'function' ? fmtStanica(sid) : sid;
                const d = fmtD(r.Datum);
                return `
                    <div class="tbl__row">
                        <div><div class="tbl__cell-main">${escapeHtml(r.KooperantName || r.KooperantID || '')}</div></div>
                        <div>
                            <div class="tbl__cell-main">${escapeHtml(r.VrstaVoca || '')}</div>
                            <div class="tbl__cell-sub">Klasa ${escapeHtml(r.Klasa || 'I')}</div>
                        </div>
                        <div><div class="tbl__cell-sub">${escapeHtml(stanicaName)}</div></div>
                        <div style="text-align:right;"><div class="tbl__num">${kg.toLocaleString('sr')}<span class="u">kg</span></div></div>
                        <div style="text-align:right;"><div class="tbl__num">${cena.toLocaleString('sr')}<span class="u">RSD</span></div></div>
                        <div><div class="tbl__cell-sub">${escapeHtml(d ? (typeof fmtDate === 'function' ? fmtDate(d) : d) : '')}</div></div>
                    </div>`;
            }).join('')}
        </div>`;
}
