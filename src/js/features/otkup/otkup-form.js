function populateVrstaDropdown() {
    const sel = document.getElementById('fldVrsta');
    sel.innerHTML = '<option value="">-- Izaberi --</option>';
    const vrsteSet = new Set();
    (stammdaten.kulture || []).forEach(k => {
        if (k.VrstaVoca && !vrsteSet.has(k.VrstaVoca)) {
            vrsteSet.add(k.VrstaVoca);
            const opt = document.createElement('option'); opt.value = k.VrstaVoca; opt.textContent = k.VrstaVoca; sel.appendChild(opt);
        }
    });
    populateKooperantDropdown();
}

function populateKooperantDropdown() {
    const sel = document.getElementById('fldKooperantManual');
    sel.innerHTML = '<option value="">-- Izaberi --</option>';
    (stammdaten.kooperanti || []).filter(k => k.StanicaID === CONFIG.OTKUPAC_ID).forEach(k => {
        const opt = document.createElement('option'); opt.value = k.KooperantID;
        opt.textContent = k.Ime + ' ' + k.Prezime + ' (' + k.KooperantID + ')'; sel.appendChild(opt);
    });
}

function onManualKooperantChange() {
    const koopID = document.getElementById('fldKooperantManual').value;
    if (!koopID) return;
    const koop = (stammdaten.kooperanti || []).find(k => k.KooperantID === koopID);
    setKooperant(koopID, koop ? (koop.Ime + ' ' + koop.Prezime) : koopID);
}

function populateParcelaDropdown(kooperantID) {
    const sel = document.getElementById('fldParcela');
    const group = document.getElementById('parcelaGroup');
    sel.innerHTML = '<option value="">-- Bez parcele --</option>';
    const parcele = (stammdaten.parcele || []).filter(p => p.KooperantID === kooperantID);
    if (parcele.length === 0) { group.style.display = 'none'; return; }
    group.style.display = 'block';
    parcele.forEach(p => { const opt = document.createElement('option'); opt.value = p.ParcelaID; opt.textContent = p.KatBroj + ' - ' + (p.Kultura || '') + ' (' + p.ParcelaID + ')'; sel.appendChild(opt); });
}

function applyDefaults() {
    const config = stammdaten.config || [];
    const dv = config.find(c => c.Parameter === 'DefaultVrsta');
    if (dv && dv.Vrednost) {
        document.getElementById('fldVrsta').value = dv.Vrednost;
        onVrstaChange();
        const ds = config.find(c => c.Parameter === 'DefaultSorta');
        if (ds && ds.Vrednost) document.getElementById('fldSorta').value = ds.Vrednost;
    }
    applyDefaultCena();
}

function applyDefaultCena() {
    const vrsta = document.getElementById('fldVrsta').value;
    if (!vrsta) return;
    const cc = (stammdaten.config || []).find(c => c.Parameter === 'Cena' + vrsta);
    if (cc && cc.Vrednost) { const f = document.getElementById('fldCena'); if (!f.value || f.value === '0') f.value = cc.Vrednost; }
}

function onVrstaChange() {
    const vrsta = document.getElementById('fldVrsta').value;
    const sel = document.getElementById('fldSorta');
    sel.innerHTML = '<option value="">-- Izaberi --</option>';
    (stammdaten.kulture || []).filter(k => k.VrstaVoca === vrsta).forEach(k => {
        if (k.SortaVoca) { const opt = document.createElement('option'); opt.value = k.SortaVoca; opt.textContent = k.SortaVoca; sel.appendChild(opt); }
    });
    applyDefaultCena();
}
