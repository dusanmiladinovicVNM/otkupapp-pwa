// ============================================================
// DISPECER STATE
// ============================================================
let dpDem = [];
let dpPlans = [];
let dpSel = null;

// status po kamionu
let dpKS = {};

// master lista kamiona za prikaz
let dpKamioni = [];

let dpKap = {};
try {
    const raw = localStorage.getItem('dpKap');
    if (raw) {
        const parsed = JSON.parse(raw);
        if (parsed && typeof parsed === 'object' && !Array.isArray(parsed)) {
            dpKap = parsed;
        }
    }
} catch (e) {
    console.error('dpKap localStorage corrupted, resetting:', e);
    dpKap = {};
    try { localStorage.removeItem('dpKap'); } catch (_) {}
}

// ============================================================
// HELPERS
// ============================================================
function dpToday() {
    const d = new Date();
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${y}-${m}-${day}`;
}

function dpSN(stanicaID) {
    if (!stanicaID) return '?';
    const s = (stammdaten.stanice || []).find(x => x.StanicaID === stanicaID);
    return s ? (s.Naziv || s.Mesto || stanicaID) : stanicaID;
}

function dpSaveKap() {
    localStorage.setItem('dpKap', JSON.stringify(dpKap || {}));
}

function dpSetK(vid, kg) {
    dpKap[vid] = parseInt(kg) || 0;
    dpSaveKap();
}

function dpSetS(vid, status, ruta) {
    dpKS[vid] = {
        status: status || 'slobodan',
        ruta: ruta || ''
    };
}

function dpCalcRuta(vid) {
    const plans = (dpPlans || [])
        .filter(p => p.VozacID === vid && (p.Status === 'planned' || p.Status === 'u_toku'));

    if (!plans.length) return '';

    const stanice = [];
    const kupci = [];

    plans.forEach(p => {
        const stanica = p.StanicaName || p.StanicaID || '';
        const kupac = p.KupacName || p.KupacID || '';

        if (stanica && !stanice.includes(stanica)) stanice.push(stanica);
        if (kupac && !kupci.includes(kupac)) kupci.push(kupac);
    });

    return [...stanice, ...kupci].join(' → ');
}



function dpGetSup() {
    const today = dpToday();
    return ((mgmtData && mgmtData.otkupiAll) || []).filter(r =>
        !(r.VozacID || r.VozaciID || '') &&
        fmtDate(r.Datum) === today
    );
}

function dpGetAsg() {
    const today = dpToday();
    return ((mgmtData && mgmtData.otkupiAll) || []).filter(r =>
        !!(r.VozacID || r.VozaciID || '') &&
        fmtDate(r.Datum) === today
    );
}

// ============================================================
// INIT / REFRESH
// ============================================================
async function dpInit() {
    dpKamioni = [];
    dpKS = {};

    (stammdaten.vozaci || []).forEach(v => {
        const vid = v.VozacID || v.vozacID || v.ID || '';
        if (!vid) return;

        if (!dpKamioni.some(x => x.id === vid)) {
            dpKamioni.push({
                id: vid,
                name: v.Naziv || v.Ime || v.Vozac || vid
            });
        }

        const kap = parseInt(v.KapacitetKG) || 0;
        if (kap > 0 && !dpKap[vid]) {
            dpKap[vid] = kap;
        }
    });

    dpSaveKap();

    await safeAsync(async () => {
        const json = await apiFetch('action=getKamionStatus');
        if (json && json.success) {
            const records = json.records || [];

            records.forEach(r => {
                const vid = r.VozacID || r.vozacID || '';
                if (!vid) return;

                dpKS[vid] = {
                    status: r.Status || r.status || 'slobodan',
                    ruta: r.Ruta || r.ruta || ''
                };

                if (!dpKamioni.some(x => x.id === vid)) {
                    dpKamioni.push({
                        id: vid,
                        name: r.VozacName || r.Naziv || vid
                    });
                }
            });
        }
    }, 'Greška pri učitavanju statusa kamiona');

    dpGetAsg().forEach(r => {
        const vid = r.VozacID || r.VozaciID || '';
        if (!vid) return;
        if (!dpKamioni.some(x => x.id === vid)) {
            dpKamioni.push({ id: vid, name: vid });
        }
    });

    Object.keys(dpKap).forEach(vid => {
        if (!dpKamioni.some(x => x.id === vid)) {
            dpKamioni.push({ id: vid, name: vid });
        }
    });

    Object.keys(dpKS).forEach(vid => {
        if (!dpKamioni.some(x => x.id === vid)) {
            dpKamioni.push({ id: vid, name: vid });
        }
    });

    dpDem = [];
    dpPlans = [];

    await safeAsync(async () => {
        const j = await apiFetch('action=getDispecer');
        if (j && j.success) {
            dpDem = j.demand || [];
            dpPlans = j.plans || [];
        }
    }, 'Greška pri učitavanju dispečera');

    const vozaciSaPlanovima = [...new Set((dpPlans || []).map(p => p.VozacID).filter(Boolean))];
    vozaciSaPlanovima.forEach(vid => {
        const ruta = dpCalcRuta(vid);
        if (!dpKS[vid]) dpKS[vid] = {};
        dpKS[vid].ruta = ruta;
        if (!dpKS[vid].status || dpKS[vid].status === 'slobodan') {
            dpKS[vid].status = 'utovar';
        }
    });
    
    dpBindDelegates();
    
    dpPD();
    dpRS();
    dpRTr();
    dpRD();
    dpRP();
    dpRK();
    dpX();

    const rt = document.getElementById('dpRT');
    if (rt) {
        rt.textContent = 'Ažurirano: ' + new Date().toLocaleTimeString('sr', {
            hour: '2-digit',
            minute: '2-digit'
        });
    }
}
// ============================================================
// DROPDOWNS
// ============================================================
function dpPD() {
    const k = document.getElementById('dpDK');
    if (k && k.options.length <= 1) {
        (stammdaten.kupci || []).forEach(x => {
            const o = document.createElement('option');
            o.value = x.KupacID;
            o.textContent = x.Naziv || x.KupacID;
            k.appendChild(o);
        });

        if (mgmtData && mgmtData.saldoKupci) {
            mgmtData.saldoKupci.forEach(x => {
                const v = x.KupacID || x.Kupac;
                if (!v || Array.from(k.options).some(o => o.value === v)) return;
                const o = document.createElement('option');
                o.value = v;
                o.textContent = x.Kupac || v;
                k.appendChild(o);
            });
        }
    }

    const vs = document.getElementById('dpDV');
    if (vs && vs.options.length <= 1) {
        const seen = new Set();
        (stammdaten.kulture || []).forEach(x => {
            if (x.VrstaVoca && !seen.has(x.VrstaVoca)) {
                seen.add(x.VrstaVoca);
                const o = document.createElement('option');
                o.value = x.VrstaVoca;
                o.textContent = x.VrstaVoca;
                vs.appendChild(o);
            }
        });
    }
}

function dpBindDelegates() {
    const sb = document.getElementById('dpSB');
    if (sb && !sb.dataset.bound) {
        sb.addEventListener('click', function (event) {
            const card = event.target.closest('[data-action="dp-select-supply"]');
            if (!card) return;
            dpTS(card.dataset.sid || '');
        });
        sb.dataset.bound = '1';
    }

    const tb = document.getElementById('dpTB');
    if (tb && !tb.dataset.bound) {
        tb.addEventListener('click', function (event) {
            const statusBtn = event.target.closest('[data-action="dp-set-status"]');
            if (statusBtn) {
                dpCS(statusBtn.dataset.vid || '', statusBtn.dataset.status || '');
                return;
            }

            const card = event.target.closest('[data-action="dp-select-truck"]');
            if (!card) return;

            if (event.target.closest('.dp-cap-input')) return;
            dpTK(card.dataset.vid || '');
        });

        tb.addEventListener('change', function (event) {
            const input = event.target.closest('.dp-cap-input');
            if (!input) return;

            dpSetK(input.dataset.vid || '', parseInt(input.value, 10) || 0);
            dpRTr();
        });

        tb.dataset.bound = '1';
    }

    const dl = document.getElementById('dpDL2');
    if (dl && !dl.dataset.bound) {
        dl.addEventListener('click', function (event) {
            const card = event.target.closest('[data-action="dp-select-demand"]');
            if (!card) return;
            dpTD(card.dataset.did || '');
        });
        dl.dataset.bound = '1';
    }

    const planList = document.getElementById('dpPlanList');
    if (planList && !planList.dataset.bound) {
        planList.addEventListener('click', function (event) {
            const statusBtn = event.target.closest('[data-action="dp-plan-status"]');
            if (statusBtn) {
                dpChgPlanSt(statusBtn.dataset.planId || '', statusBtn.dataset.status || '');
                return;
            }

            const removeBtn = event.target.closest('[data-action="dp-remove-plan"]');
            if (removeBtn) {
                dpRmPlan(removeBtn.dataset.planId || '');
            }
        });

        planList.dataset.bound = '1';
    }
}

// ============================================================
// SUPPLY
// ============================================================
function dpRS() {
    const b = document.getElementById('dpSB');
    if (!b) return;

    const sup = dpGetSup();
    const st = document.getElementById('dpST');
    if (st) {
        const totalKg = sup.reduce((s, r) => s + (parseFloat(r.Kolicina) || 0), 0);
        st.textContent = totalKg.toLocaleString('sr') + ' kg';
    }
    const g = {};

    sup.forEach(r => {
        const s = r.OtkupacID || (r._sheetName || '').replace('OTK-', '') || '?';
        if (!g[s]) g[s] = { kg: 0, n: 0, rows: [] };
        g[s].kg += parseFloat(r.Kolicina) || 0;
        g[s].n++;
        g[s].rows.push(r);
    });

    const ids = Object.keys(g).sort((a, b) => (g[b].kg || 0) - (g[a].kg || 0));

    if (!ids.length) {
        b.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:20px;">Nema neraspoređene robe za danas</p>';
        return;
    }

    b.innerHTML = ids.map(sid => {
        const x = g[sid];
        const isSel = dpSel && dpSel.step >= 2 && dpSel.sid === sid;
        return `
            <div class="dp-card sup${isSel ? ' sel' : ''}" data-action="dp-select-supply" data-sid="${escapeHtml(String(sid || ''))}">
                <div style="display:flex;justify-content:space-between;align-items:center;">
                    <span class="dp-icon-line" style="font-weight:700;">
                        ${agIcon('package', '16px')} ${escapeHtml(dpSN(sid))}
                    </span>
                    <span style="font-weight:700;">${x.kg.toLocaleString('sr')} kg</span>
                </div>
                <div style="font-size:12px;color:var(--text-muted);margin-top:4px;">
                    ${x.n} otkupa
                </div>
            </div>
        `;
    }).join('');
}

// ============================================================
// TRANSPORT
// ============================================================
function dpRTr() {
    const b = document.getElementById('dpTB');
    if (!b) return;

    const asg = dpGetAsg();
    const vm = {};

    // 1) svi poznati kamioni
    (dpKamioni || []).forEach(v => {
        const vid = v.id || v.VozacID || v.vozacID || '';
        if (!vid) return;

        vm[vid] = {
            kg: 0,              // stvarno dodeljeno kroz otkupe
            plannedKg: 0,       // planirano kroz dispečer planove
            n: 0,               // broj stvarno dodeljenih otkupa
            st: new Set(),      // stanice iz realnih otkupa i planova
            name: v.name || v.Naziv || vid,
            plans: []           // svi aktivni planovi za kamion
        };
    });

    // 2) stvarno assigned otkupi
    asg.forEach(r => {
        const vid = r.VozacID || r.VozaciID || '';
        if (!vid) return;

        if (!vm[vid]) {
            vm[vid] = {
                kg: 0,
                plannedKg: 0,
                n: 0,
                st: new Set(),
                name: vid,
                plans: []
            };
        }

        vm[vid].kg += parseFloat(r.Kolicina) || 0;
        vm[vid].n++;

        const sid = r.OtkupacID || (r._sheetName || '').replace('OTK-', '');
        if (sid) vm[vid].st.add(sid);
    });

    // 3) aktivni planovi
    (dpPlans || []).forEach(p => {
        const vid = p.VozacID || '';
        if (!vid) return;
        if (p.Status !== 'planned' && p.Status !== 'u_toku') return;

        if (!vm[vid]) {
            vm[vid] = {
                kg: 0,
                plannedKg: 0,
                n: 0,
                st: new Set(),
                name: p.VozacName || vid,
                plans: []
            };
        }

        const pKg = parseFloat(p.PlannedKg) || 0;
        vm[vid].plannedKg += pKg;
        vm[vid].plans.push(p);

        const sid = p.StanicaID || '';
        if (sid) vm[vid].st.add(sid);
    });

    // 4) fallback iz kapaciteta
    Object.keys(dpKap).forEach(vid => {
        if (!vm[vid]) {
            vm[vid] = {
                kg: 0,
                plannedKg: 0,
                n: 0,
                st: new Set(),
                name: vid,
                plans: []
            };
        }
    });

    // 5) fallback iz statusa
    Object.keys(dpKS).forEach(vid => {
        if (!vm[vid]) {
            vm[vid] = {
                kg: 0,
                plannedKg: 0,
                n: 0,
                st: new Set(),
                name: vid,
                plans: []
            };
        }
    });

    // 6) sortiranje — najviše ukupno opterećeni prvi
    const ids = Object.keys(vm).sort((a, b) => {
        const ak = (vm[a].kg || 0) + (vm[a].plannedKg || 0);
        const bk = (vm[b].kg || 0) + (vm[b].plannedKg || 0);
        if (bk !== ak) return bk - ak;
        return (vm[a].name || a).localeCompare(vm[b].name || b);
    });

    const tt = document.getElementById('dpTT');
    if (tt) tt.textContent = ids.length;

    if (!ids.length) {
        b.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:20px;">Nema kamiona za prikaz</p>';
        return;
    }

    const sl = {
        slobodan: 'Slobodan',
        utovar: 'Utovar',
        naputu: 'Na putu',
        istovar: 'Istovar'
    };

    b.innerHTML = ids.map(vid => {
        const x = vm[vid];
        const cap = parseInt(dpKap[vid]) || 0;

        const realKg = x.kg || 0;
        const plannedKg = x.plannedKg || 0;
        const loadKg = realKg + plannedKg;

        const pct = cap > 0 ? Math.min(100, Math.round((loadKg / cap) * 100)) : 0;
        const freeKg = cap > 0 ? Math.max(0, cap - loadKg) : 0;

        const bc =
            pct >= 95 ? 'var(--danger)' :
            pct >= 75 ? 'var(--accent)' :
            'var(--success)';

        const isSel = dpSel && dpSel.step >= 1 && dpSel.vid === vid;
        const st = (dpKS[vid] || {}).status || 'slobodan';
        const ruta = (dpKS[vid] || {}).ruta || '';

        // sve stanice kroz planove + assigned
        const stationNames = [...(x.st || [])].map(s => dpSN(s));

        // hladnjače iz aktivnih planova
        const planKupci = [...new Set(
            (x.plans || [])
                .map(p => p.KupacName || p.KupacID || '')
                .filter(Boolean)
        )];

        let rutaText = '';
        if (stationNames.length && planKupci.length) {
            rutaText = stationNames.join(' → ') + ' → ' + planKupci.join(', ');
        } else if (stationNames.length && ruta) {
            rutaText = ruta;
        } else if (ruta) {
            rutaText = ruta;
        }

        const planHtml = (x.plans || []).map(p => {
            const pKg = parseInt(p.PlannedKg || 0) || 0;
            return `
                <div style="font-size:11px;margin-top:3px;color:var(--success);font-weight:600;">
                    ${agIcon('clipboard', '13px')} Plan: ${escapeHtml(p.StanicaName || p.StanicaID || '?')} → ${escapeHtml(p.KupacName || p.KupacID || '?')} (${pKg.toLocaleString('sr')} kg)
                </div>
            `;
        }).join('');

        return `
            <div class="dp-card trn${isSel ? ' sel' : ''}" data-action="dp-select-truck" data-vid="${escapeHtml(String(vid || ''))}">
                <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:2px;">
                    <span class="dp-icon-line" style="font-weight:700;font-size:14px;">
                        ${agIcon('truck', '16px')} ${escapeHtml(x.name || vid)}
                    </span>
                    <span style="font-size:14px;font-weight:700;">${loadKg.toLocaleString('sr')} kg</span>
                </div>

                <div style="font-size:12px;color:var(--text-muted);margin-top:2px;">
                    Kap: ${
                        cap > 0
                            ? cap.toLocaleString('sr') + ' kg'
                            : `<input type="number"
                                      inputmode="numeric"
                                      placeholder="kg"
                                      class="dp-cap-input"
                                      data-vid="${escapeHtml(String(vid || ''))}"
                                      style="width:70px;padding:2px 4px;font-size:11px;border:1px solid var(--border);border-radius:4px;">`
                    }
                    · Popunjeno: <strong>${pct}%</strong>
                    ${cap > 0 ? ` · Slobodno: ${freeKg.toLocaleString('sr')} kg` : ''}
                </div>

                ${cap > 0 ? `
                    <div class="dp-bar" style="margin-top:6px;">
                        <div class="dp-bf" style="width:${pct}%;background:${bc};"></div>
                    </div>
                ` : ''}

                <div style="font-size:11px;color:var(--text-muted);margin-top:4px;">
                    ${realKg > 0 ? `Realno: ${realKg.toLocaleString('sr')} kg` : 'Realno: 0 kg'}
                    ${plannedKg > 0 ? ` · Planirano: ${plannedKg.toLocaleString('sr')} kg` : ''}
                    · ${x.n} otk.
                </div>

                ${rutaText ? `
                    <div style="font-size:11px;color:var(--text-muted);margin-top:4px;">
                        Ruta: ${escapeHtml(rutaText)}
                    </div>
                ` : ''}

                ${planHtml}

                <div style="display:flex;justify-content:space-between;align-items:center;margin-top:6px;">
                    <span class="dp-badge ${st}">${escapeHtml(sl[st] || st)}</span>
                </div>

                <div class="dp-stb">
                    ${['slobodan', 'utovar', 'naputu', 'istovar']
                        .map(s => `<button type="button" class="${st === s ? 'on' : ''}" data-action="dp-set-status" data-vid="${escapeHtml(String(vid || ''))}" data-status="${s}">${sl[s]}</button>`)
                        .join('')}
                </div>
            </div>
        `;
    }).join('');
}
// ============================================================
// DEMAND
// ============================================================
function dpRD() {
    const l = document.getElementById('dpDL2');
    const t = document.getElementById('dpDT');
    if (!l || !t) return;

    const tot = dpDem.reduce((s, d) => s + (parseInt(d.Kg) || 0), 0);
    t.textContent = tot.toLocaleString('sr') + ' kg';

    if (!dpDem.length) {
        l.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:20px;">Nema demand-a za danas</p>';
        return;
    }

    // saberi aktivne planove po DemandID
    const plannedByDemand = {};
    (dpPlans || []).forEach(p => {
        if (p.Status !== 'planned' && p.Status !== 'u_toku') return;

        const did = p.DemandID || '';
        if (!did) return;

        plannedByDemand[did] = (plannedByDemand[did] || 0) + (parseInt(p.PlannedKg) || 0);
    });

    l.innerHTML = dpDem.map(d => {
        const did = d.DemandID || d.demandID || '';
        const isSel = dpSel && dpSel.step >= 3 && dpSel.did === did;

        const kup = d.KupacName || d.KupacID || '?';
        const trazeno = parseInt(d.Kg) || 0;
        const primljeno = parseInt(d.Primljeno) || 0;
        const planirano = plannedByDemand[did] || 0;

        // koliko još nije ni planirano ni primljeno
        const preostalo = Math.max(0, trazeno - planirano - primljeno);

        // ukupno "pokriveno"
        const pokriveno = Math.min(trazeno, planirano + primljeno);
        const pct = trazeno > 0 ? Math.min(100, Math.round((pokriveno / trazeno) * 100)) : 0;

        const barColor =
            pct >= 100 ? 'var(--success)' :
            pct >= 70 ? 'var(--accent)' :
            '#1565c0';

        return `
            <div class="dp-card dem${isSel ? ' sel' : ''}" data-action="dp-select-demand" data-did="${escapeHtml(String(did || ''))}">
                <div style="display:flex;justify-content:space-between;align-items:center;">
                    <strong class="dp-icon-line">
                        ${agIcon('factory', '16px')} ${escapeHtml(kup)}
                    </strong>
                    <strong>${trazeno.toLocaleString('sr')} kg</strong>
                </div>

                <div style="font-size:12px;color:var(--text-muted);margin-top:4px;">
                    ${escapeHtml(d.Vrsta || '')} ${escapeHtml(d.Klasa || '')}
                </div>

                <div class="dp-bar" style="margin-top:8px;">
                    <div class="dp-bf" style="width:${pct}%;background:${barColor};"></div>
                </div>

                <div style="font-size:11px;color:var(--text-muted);margin-top:6px;line-height:1.5;">
                    <div>Planirano: <strong style="color:#1565c0;">${planirano.toLocaleString('sr')} kg</strong></div>
                    <div>Primljeno: <strong style="color:var(--success);">${primljeno.toLocaleString('sr')} kg</strong></div>
                    <div>Preostalo: <strong style="color:${preostalo > 0 ? 'var(--danger)' : 'var(--success)'};">${preostalo.toLocaleString('sr')} kg</strong></div>
                </div>
            </div>
        `;
    }).join('');
}
// ============================================================
// PLANOVI
// ============================================================
function dpRP() {
    const b = document.getElementById('dpPlanList');
    if (!b) return;

    if (!dpPlans.length) {
        b.innerHTML = '';
        return;
    }

    b.innerHTML = dpPlans.map(p => `
        <div class="dp-plan-item">
            <div>
                <div class="dp-pi-route dp-icon-line">
                    ${agIcon('truck', '15px')} ${escapeHtml(p.VozacID)} · ${escapeHtml(p.StanicaName || p.StanicaID || '?')} → ${escapeHtml(p.KupacName || p.KupacID || '?')}
                </div>
                <div style="font-size:11px;color:var(--text-muted);margin-top:2px;">
                    ${parseInt(p.PlannedKg || 0).toLocaleString('sr')} kg · status: ${escapeHtml(p.Status || 'planned')}
                </div>
            </div>
            <div style="display:flex;gap:6px;">
                <button type="button" data-action="dp-plan-status" data-plan-id="${escapeHtml(String(p.PlanID || ''))}" data-status="u_toku" title="U toku" aria-label="U toku">
                    ${agIcon('arrow-right', '14px')}
                </button>
                <button type="button" data-action="dp-plan-status" data-plan-id="${escapeHtml(String(p.PlanID || ''))}" data-status="zavrseno" title="Završeno" aria-label="Završeno">
                    ${agIcon('check', '14px')}
                </button>
                <button type="button" data-action="dp-remove-plan" data-plan-id="${escapeHtml(String(p.PlanID || ''))}" title="Obriši" aria-label="Obriši">
                    ${agIcon('x', '14px')}
                </button>
            </div>
        </div>
    `).join('');
}

// ============================================================
// KPI
// ============================================================
function dpRK() {
    const k1 = document.getElementById('dpK1');
    const k2 = document.getElementById('dpK2');
    const k3 = document.getElementById('dpK3');
    const k4 = document.getElementById('dpK4');
    if (!k1 || !k2 || !k3 || !k4) return;

    k1.textContent = dpGetSup()
        .reduce((s, r) => s + (parseFloat(r.Kolicina) || 0), 0)
        .toLocaleString('sr');

    // ovde sada brojimo master listu kamiona, ne samo assigned
    const vz = new Set();
    (dpKamioni || []).forEach(v => vz.add(v.id || v.VozacID || v.vozacID));
    Object.keys(dpKap).forEach(v => vz.add(v));
    Object.keys(dpKS).forEach(v => vz.add(v));

    k2.textContent = vz.size;
    k3.textContent = dpDem.reduce((s, d) => s + (parseInt(d.Kg) || 0), 0).toLocaleString('sr');
    k4.textContent = dpPlans.filter(p => p.Status === 'planned' || p.Status === 'u_toku').length;
}

// ============================================================
// TAP TO PLAN
// ============================================================
function dpTK(vid) {
    if (dpSel && dpSel.vid === vid) {
        dpX();
        return;
    }
    dpSel = { step: 1, vid };
    dpBN('🚛 ' + vid + ' izabran', 'Korak 2: tapnite STANICU odakle treba pokupiti robu');
    dpHL();
}

function dpTS(sid) {
    if (!dpSel || dpSel.step < 1) {
        showToast('Prvo izaberite kamion', 'info');
        return;
    }
    if (dpSel.sid === sid) {
        dpSel.step = 1;
        dpSel.sid = null;
        dpBN('🚛 ' + dpSel.vid, 'Korak 2: tapnite stanicu');
        dpHL();
        return;
    }

    dpSel.step = 2;
    dpSel.sid = sid;

    const kg = dpGetSup()
        .filter(r => (r.OtkupacID || (r._sheetName || '').replace('OTK-', '')) === sid)
        .reduce((s, r) => s + (parseFloat(r.Kolicina) || 0), 0);

    dpBN(
        '🚛 ' + escapeHtml(dpSel.vid) + ' → 📦 ' + escapeHtml(dpSN(sid)) + ' (' + kg.toLocaleString('sr') + ' kg)',
        'Korak 3: tapnite HLADNJAČU gde ide roba'
    );
    dpHL();
}

function dpTD(did) {
    if (!dpSel || dpSel.step < 2) {
        showToast(dpSel && dpSel.step >= 1 ? 'Izaberite stanicu' : 'Prvo kamion, pa stanicu', 'info');
        return;
    }
    if (dpSel.did === did) {
        dpSel.step = 2;
        dpSel.did = null;
        dpBN('🚛 ' + dpSel.vid + ' → 📦 ' + dpSN(dpSel.sid), 'Korak 3: tapnite hladnjaču');
        dpHL();
        return;
    }

    dpSel.step = 3;
    dpSel.did = did;

    const d = dpDem.find(x => x.DemandID === did);
    const kg = dpGetSup()
        .filter(r => (r.OtkupacID || (r._sheetName || '').replace('OTK-', '')) === dpSel.sid)
        .reduce((s, r) => s + (parseFloat(r.Kolicina) || 0), 0);

    dpBN(
        '🚛 ' + escapeHtml(dpSel.vid) + ' → 📦 ' + escapeHtml(dpSN(dpSel.sid)) + ' (' + kg.toLocaleString('sr') + ' kg) → 🏭 ' + escapeHtml(d ? d.KupacName : '?'),
        'Tapnite SAČUVAJ PLAN'
    );
    dpHL();
}

function dpBN(t, s) {
    document.getElementById('dpBt').textContent = t;
    document.getElementById('dpBs').textContent = s;
    document.getElementById('dpBnr').classList.add('active');
}

function dpX() {
    dpSel = null;
    document.getElementById('dpBnr').classList.remove('active');
    dpHL();
}

function dpHL() {
    dpRS();
    dpRTr();
    dpRD();
}

// ============================================================
// SAVE PLAN
// ============================================================
async function dpOK() {
    if (!dpSel || dpSel.step < 3) {
        showToast('Završite sva 3 koraka', 'error');
        return;
    }

    const vid = dpSel.vid;
    const sid = dpSel.sid;
    const d = dpDem.find(x => x.DemandID === dpSel.did);
    const kupN = d ? (d.KupacName || d.KupacID) : '?';
    const kupID = d ? (d.KupacID || '') : '';

    const kg = dpGetSup()
        .filter(r => (r.OtkupacID || (r._sheetName || '').replace('OTK-', '')) === sid)
        .reduce((s, r) => s + (parseFloat(r.Kolicina) || 0), 0);

    showToast('Čuvanje plana.', 'info');

    const json = await safeAsync(async () => {
        return await apiPost('saveDispecer', {
            demandID: dpSel.did,
            vozacID: vid,
            stanicaID: sid,
            stanicaName: dpSN(sid),
            kupacID: kupID,
            kupacName: kupN,
            plannedKg: Math.round(kg)
        });
    }, 'Greška pri čuvanju plana');

    if (!json) { dpX(); return; }

    if (json.success) {
        const newPlan = {
            PlanID: json.planID,
            DemandID: dpSel.did,
            VozacID: vid,
            StanicaID: sid,
            StanicaName: dpSN(sid),
            KupacID: kupID,
            KupacName: kupN,
            PlannedKg: Math.round(kg),
            Status: 'planned'
        };

        dpPlans.push(newPlan);

        const ruta = dpCalcRuta(vid);
        dpSetS(vid, 'utovar', ruta);

        // Fire-and-forget status update
        apiPost('updateKamionStatus', {
            vozacID: vid,
            status: 'utovar',
            ruta: ruta
        }).catch(() => {});

        if (!dpKamioni.some(x => x.id === vid)) {
            dpKamioni.push({ id: vid, name: vid });
        }

        showToast('📋 Plan: ' + ruta, 'success');
    } else {
        showToast(json.error || 'Greška', 'error');
    }

    dpX();
    dpRS();
    dpRTr();
    dpRP();
    dpRK();
}
// ============================================================
// PLAN STATUS CHANGE
// ============================================================
async function dpChgPlanSt(planID, newStatus) {
    const json = await safeAsync(async () => {
        return await apiPost('updateDispecer', {
            planID: planID,
            status: newStatus
        });
    }, 'Greška pri izmeni statusa plana');

    if (!json) return;

    const p = dpPlans.find(x => x.PlanID === planID);
    const vid = p ? p.VozacID : '';

    if (p) p.Status = newStatus;
    if (newStatus === 'zavrseno') {
        dpPlans = dpPlans.filter(x => x.PlanID !== planID);
    }

    if (vid) {
        const ruta = dpCalcRuta(vid);
        const status = ruta
            ? (newStatus === 'zavrseno' ? ((dpKS[vid] && dpKS[vid].status) || 'utovar') : newStatus)
            : 'slobodan';

        dpSetS(vid, status, ruta);

        apiPost('updateKamionStatus', {
            vozacID: vid,
            status: status,
            ruta: ruta
        }).catch(() => {});
    }

    dpRP();
    dpRK();
    dpRS();
    dpRTr();

    showToast(newStatus === 'zavrseno' ? 'Plan završen' : 'Plan u toku', 'success');
}

async function dpRmPlan(planID) {
    const plan = dpPlans.find(x => x.PlanID === planID);
    const vid = plan ? plan.VozacID : '';

    const json = await safeAsync(async () => {
        return await apiPost('removeDispecer', { planID: planID });
    }, 'Greška pri brisanju plana');

    if (!json) return;

    dpPlans = dpPlans.filter(x => x.PlanID !== planID);

    if (vid) {
        const ruta = dpCalcRuta(vid);
        const status = ruta ? ((dpKS[vid] && dpKS[vid].status) || 'utovar') : 'slobodan';

        dpSetS(vid, status, ruta);

        apiPost('updateKamionStatus', {
            vozacID: vid,
            status: status,
            ruta: ruta
        }).catch(() => {});
    }

    dpRP();
    dpRK();
    dpRS();
    dpRTr();

    showToast('Plan obrisan', 'info');
}
    

// ============================================================
// KAMION STATUS
// ============================================================

async function dpCS(vid, st) {
    dpKS[vid] = {
        status: st,
        ruta: (dpKS[vid] || {}).ruta || ''
    };

    if (!dpKamioni.some(x => x.id === vid)) {
        dpKamioni.push({ id: vid, name: vid });
    }

    dpRTr();

    await safeAsync(async () => {
        return await apiPost('updateKamionStatus', {
            vozacID: vid,
            status: st,
            ruta: (dpKS[vid] || {}).ruta || ''
        });
    }, 'Greška pri čuvanju statusa kamiona');
}

async function dpAD() {
    const kupacID = document.getElementById('dpDK').value;
    const kg = parseInt(document.getElementById('dpDG').value) || 0;
    const vrsta = document.getElementById('dpDV').value || '';
    const klasa = document.getElementById('dpDL').value || '';

    if (!kupacID) {
        showToast('Izaberite kupca', 'error');
        return;
    }
    if (kg <= 0) {
        showToast('Unesite kg', 'error');
        return;
    }

    const kupacName =
        document.getElementById('dpDK').selectedOptions[0]?.textContent || kupacID;

    const json = await safeAsync(async () => {
        return await apiPost('saveWarRoomDemand', {
            kupacID: kupacID,
            kupacName: kupacName,
            kg: kg,
            vrsta: vrsta,
            klasa: klasa
        });
    }, 'Greška pri čuvanju zahteva');

    if (!json) return;

    if (!json.success) {
        showToast(json.error || 'Greška pri čuvanju', 'error');
        return;
    }

    dpDem.push({
        DemandID: json.demandID,
        KupacID: kupacID,
        KupacName: kupacName,
        Kg: kg,
        Vrsta: vrsta,
        Klasa: klasa,
        Primljeno: 0
    });

    document.getElementById('dpDG').value = '';
    document.getElementById('dpDV').value = '';
    document.getElementById('dpDL').value = '';

    dpRD();
    dpRK();

    showToast('Zahtev dodat', 'success');
}

async function loadDispecer() { await dpInit(); }
