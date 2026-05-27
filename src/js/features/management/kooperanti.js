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
    if (!koopID) {
        const panel = document.getElementById('mgmtKoopDetailPanel');
        if (panel) panel.innerHTML = '<div class="partner-detail-empty">Izaberi kooperanta iz liste sa leve strane</div>';
        return;
    }
    await onMgmtKooperantChangeDirect(koopID);
}

let _mgmtKoopSaldoCache = [];

async function loadMgmtKoopSaldo() {
    if (typeof ensureMgmtSection === 'function') {
        await ensureMgmtSection('kartice', 'getMgmtKartice');
    }

    const kartice = (mgmtData && Array.isArray(mgmtData.kartice))
        ? mgmtData.kartice
        : [];

    const totals = kartice.filter(r => r.Opis === 'UKUPNO');
    _mgmtKoopSaldoCache = [...totals].sort((a, b) => (parseFloat(b.Saldo)||0) - (parseFloat(a.Saldo)||0));

    // Update segment counter
    const countEl = document.getElementById('mgmtKoopCount');
    if (countEl) countEl.textContent = _mgmtKoopSaldoCache.length;

    renderMgmtKoopSaldoList('');
}

function renderMgmtKoopSaldoList(query) {
    const list = document.getElementById('mgmtKoopSaldoList');
    if (!list) return;

    const q = String(query || '').trim().toLowerCase();
    const filtered = _mgmtKoopSaldoCache.filter(r => {
        if (!q) return true;
        const koop = (stammdaten.kooperanti || []).find(k => k.KooperantID === r.KooperantID);
        const name = koop ? (koop.Ime + ' ' + koop.Prezime) : (r.KooperantID || '');
        return name.toLowerCase().includes(q) || String(r.KooperantID || '').toLowerCase().includes(q);
    });

    list.innerHTML = `
        <div class="partner-list">
            <div class="partner-list__head">
                <span class="partner-list__title">Po saldu — opadajuće</span>
                <div class="partner-list__search">
                    <input type="search" id="mgmtKoopSearch" placeholder="Pretraži kooperanta…" value="${escapeHtml(query || '')}">
                </div>
            </div>
            ${filtered.length === 0
                ? '<div class="partner-list__empty">Nema rezultata</div>'
                : filtered.map(r => {
                    const koop = (stammdaten.kooperanti || []).find(k => k.KooperantID === r.KooperantID);
                    const name = koop ? koop.Ime + ' ' + koop.Prezime : r.KooperantID;
                    const saldo = parseFloat(r.Saldo) || 0;
                    const initials = name.split(' ').filter(Boolean).map(p => p[0]).join('').slice(0,2).toUpperCase();
                    const saldoMod = saldo > 0 ? 'warn' : saldo < 0 ? 'ok' : 'neutral';
                    const meta = koop && (koop.Mesto || koop.StanicaID)
                        ? `${escapeHtml(r.KooperantID || '')} · ${escapeHtml(koop.Mesto || koop.StanicaID || '')}`
                        : escapeHtml(r.KooperantID || '');
                    return `
                        <div class="partner" data-koop-id="${escapeHtml(r.KooperantID || '')}" data-action="mgmt-koop-select">
                            <div class="partner__avatar">${escapeHtml(initials)}</div>
                            <div>
                                <div class="partner__name">${escapeHtml(name)}</div>
                                <div class="partner__meta">${meta}</div>
                            </div>
                            <div class="partner__saldo partner__saldo--${saldoMod}">
                                ${saldo > 0 ? '+' : ''}${saldo.toLocaleString('sr')}<span class="u">RSD</span>
                            </div>
                        </div>
                    `;
                }).join('')}
        </div>
    `;

    const searchEl = document.getElementById('mgmtKoopSearch');
    if (searchEl) {
        searchEl.addEventListener('input', (e) => renderMgmtKoopSaldoList(e.target.value));
        // Restore focus + caret if user was typing
        if (q) {
            searchEl.focus();
            const len = searchEl.value.length;
            try { searchEl.setSelectionRange(len, len); } catch (_) {}
        }
    }
}

async function loadMgmtKoopPregled() {
    if (typeof ensureMgmtSection === 'function') {
        await ensureMgmtSection('saldoOMDetail', 'getMgmtSaldoOMDetail');
    }
    
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

// Click delegate: selecting kooperant from saldo list loads their kartica
(function bindKoopSaldoDelegate() {
    document.addEventListener('click', function(e) {
        const item = e.target.closest('[data-action="mgmt-koop-select"]');
        if (!item) return;
        const koopID = item.dataset.koopId;
        if (!koopID) return;

        // Mark active
        document.querySelectorAll('[data-action="mgmt-koop-select"]').forEach(el => el.classList.remove('is-active'));
        item.classList.add('is-active');

        // Load kartica directly (bypass dropdown on desktop split view)
        onMgmtKooperantChangeDirect(koopID);
    });
})();

async function onMgmtKooperantChangeDirect(koopID) {
    const panel = document.getElementById('mgmtKoopDetailPanel');
    if (!panel) return;

    panel.innerHTML = '<div class="partner-detail-empty" style="padding:40px;text-align:center;color:var(--text-muted);">Učitavanje…</div>';

    // --- basic kooperant info ---
    const koop = (stammdaten && stammdaten.kooperanti || []).find(k => k.KooperantID === koopID);
    const name = koop ? koop.Ime + ' ' + koop.Prezime : koopID;

    // --- kartice ---
    if (window.mgmtData && !window.mgmtData.kartice) {
        if (typeof ensureMgmtSection === 'function') {
            await ensureMgmtSection('kartice', 'getMgmtKartice');
        }
    }
    let records = [];
    if (window.mgmtData && window.mgmtData.kartice) {
        records = window.mgmtData.kartice.filter(r => r.KooperantID === koopID && r.Opis !== 'UKUPNO');
    } else {
        const json = await safeAsync(async () => {
            return await apiFetch('action=getMgmtKartica&kooperantID=' + encodeURIComponent(koopID));
        }, 'Greška pri učitavanju kartice kooperanta');
        if (json && json.success && json.records) {
            records = json.records.filter(r => r.Opis !== 'UKUPNO');
        }
    }

    // --- saldo from UKUPNO row ---
    const ukupno = (window.mgmtData && window.mgmtData.kartice || []).find(r => r.KooperantID === koopID && r.Opis === 'UKUPNO');
    const saldo = parseFloat(ukupno && ukupno.Saldo || 0);

    // --- predato / agro from saldoOMDetail ---
    const omDetail = (window.mgmtData && window.mgmtData.saldoOMDetail) || [];
    const koopOM = omDetail.filter(r => r.KooperantID === koopID);
    const predatoKg  = koopOM.reduce((s, r) => s + (parseFloat(r.Kolicina) || 0), 0);
    const agroRSD    = koopOM.reduce((s, r) => s + (parseFloat(r.AgroZaduzenje) || 0), 0);

    // --- parcele info from stammdaten ---
    const parcele = (stammdaten && stammdaten.parcele || []).filter(p => p.KooperantID === koopID);
    const totalHa  = parcele.reduce((s, p) => s + (parseFloat(p.Povrsina || p.Ha || 0)), 0);

    // --- meta line ---
    const location = koop && (koop.Mesto || koop.Selo) ? `${koop.Mesto || koop.Selo}` : '';
    const phone    = koop && koop.Telefon ? koop.Telefon : '';
    const saldoSign = saldo >= 0 ? '+' : '';
    const saldoColor = saldo > 0 ? 'var(--color-warning)' : saldo < 0 ? 'var(--color-success)' : 'var(--text-muted)';

    // --- transaction rows (latest 10 shown, link to full) ---
    const PREVIEW = 10;
    const txnHtml = records.length === 0
        ? '<div class="partner-detail-empty">Nema transakcija</div>'
        : records.slice(0, PREVIEW).map(r => {
            const z  = parseFloat(r.Zaduzenje)  || 0;
            const ra = parseFloat(r.Razduzenje) || 0;
            const amt = ra > 0 ? ra : -z;
            const amtFmt = (ra > 0 ? '+' : '') + amt.toLocaleString('sr');
            const amtColor = ra > 0 ? 'var(--color-success)' : 'var(--color-warning)';
            const shortDate = mgmtShortDate(r.Datum);
            return `
                <div class="trans">
                    <div class="trans__date">
                        <strong>${escapeHtml(shortDate)}</strong>
                        ${escapeHtml(r.Vreme || '')}
                    </div>
                    <div>
                        <div class="trans__lbl">${escapeHtml(r.Opis || r.BrojDok || '')}</div>
                        <div class="trans__sub">${escapeHtml(r.BrojDok || '')}</div>
                    </div>
                    <div style="font-family:var(--font-display);font-size:15px;font-weight:700;color:${amtColor};white-space:nowrap;">
                        ${escapeHtml(amtFmt)}<span style="font-size:10px;color:var(--text-muted);font-weight:500;"> RSD</span>
                    </div>
                </div>`;
        }).join('');

    const moreCount = records.length > PREVIEW ? records.length - PREVIEW : 0;

    panel.innerHTML = `
        <div class="partner-detail">
            <div class="partner-detail__head">
                <div class="partner-detail__eyebrow">Kooperant · ${escapeHtml(koopID)}</div>
                <div class="partner-detail__title">${escapeHtml(name)}</div>
                ${(location || phone || parcele.length) ? `
                <div class="partner-detail__meta">
                    ${location ? '<span>📍 ' + escapeHtml(location) + '</span>' : ''}
                    ${phone ? '<span>📞 ' + escapeHtml(phone) + '</span>' : ''}
                    ${parcele.length ? '<span>🌿 ' + parcele.length + ' parcele' + (totalHa > 0 ? ' · ' + totalHa.toLocaleString('sr', {maximumFractionDigits:1}) + ' ha' : '') + '</span>' : ''}
                </div>` : ''}
            </div>
            <div class="partner-detail__stats">
                <div class="partner-detail__stat">
                    <div class="partner-detail__stat-lbl">Saldo</div>
                    <div class="partner-detail__stat-val" style="color:${saldoColor};">
                        ${saldoSign}${saldo.toLocaleString('sr')}<span class="u">RSD</span>
                    </div>
                </div>
                <div class="partner-detail__stat">
                    <div class="partner-detail__stat-lbl">Predato sezona</div>
                    <div class="partner-detail__stat-val">
                        ${predatoKg.toLocaleString('sr')}<span class="u">kg</span>
                    </div>
                </div>
                <div class="partner-detail__stat">
                    <div class="partner-detail__stat-lbl">Agrohemija</div>
                    <div class="partner-detail__stat-val">
                        ${agroRSD.toLocaleString('sr')}<span class="u">RSD</span>
                    </div>
                </div>
            </div>
            <div class="partner-detail__body">
                <div class="partner-detail__txn-head">
                    <span>Poslednje transakcije</span>
                    <div class="partner-detail__txn-actions">
                        <button class="partner-detail__action" type="button" disabled title="Uskoro">Pošalji izvod</button>
                        <button class="partner-detail__action partner-detail__action--primary" type="button" disabled title="Uskoro">Pripremi isplatu →</button>
                    </div>
                </div>
                ${txnHtml}
                ${moreCount > 0 ? `
                <div class="partner-detail__more">
                    Cela kartica (${records.length} transakcija) →
                </div>` : ''}
            </div>
        </div>`;
}

// Helper: month number (1-based string) → short Serbian name
function mgmtMonthShort(mm) {
    const m = ['', 'jan', 'feb', 'mar', 'apr', 'maj', 'jun', 'jul', 'avg', 'sep', 'okt', 'nov', 'dec'];
    return m[parseInt(mm, 10)] || mm;
}

// Helper: any date input → "12. avg" (handles ISO, dotted, Date)
function mgmtShortDate(val) {
    if (!val) return '';
    const s = String(val);
    // ISO "YYYY-MM-DD..."
    const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (iso) return parseInt(iso[3], 10) + '. ' + mgmtMonthShort(iso[2]);
    // Dotted "DD.MM.YYYY" or "D.M.YYYY"
    const dot = s.match(/^(\d{1,2})\.(\d{1,2})\./);
    if (dot) return parseInt(dot[1], 10) + '. ' + mgmtMonthShort(dot[2].padStart(2, '0'));
    return s;
}
