// ============================================================
// KOOPERANT: PARCELE
// ============================================================
let parcelMapInstance = null;

const kooperantParcelStyle = {
    color: '#ffd60a',
    weight: 3,
    opacity: 1,
    fillColor: '#ffd60a',
    fillOpacity: 0.18
};

const kooperantSelectedParcelStyle = {
    color: '#ff2d55',
    weight: 4,
    opacity: 1,
    fillColor: '#ff2d55',
    fillOpacity: 0.22
};

let activeKooperantParcelLayer = null;

function resetKooperantParcelHighlight() {
    if (activeKooperantParcelLayer && activeKooperantParcelLayer.setStyle) {
        activeKooperantParcelLayer.setStyle(kooperantParcelStyle);
    }
    activeKooperantParcelLayer = null;
}

function highlightKooperantParcelLayer(layer) {
    resetKooperantParcelHighlight();
    if (layer && layer.setStyle) {
        layer.setStyle(kooperantSelectedParcelStyle);
        activeKooperantParcelLayer = layer;
    }
}

function buildKooperantParcelPopup(p) {
    return `
        <div>
            <div style="font-size:18px;font-weight:700;margin-bottom:6px;">
                ${p.KatBroj || p.ParcelaID}
            </div>
            <div><b>Kultura:</b> ${p.Kultura || '-'}</div>
            <div><b>Površina:</b> ${p.PovrsinaHa || '?'} ha</div>
            <div><b>KO:</b> ${p.KatOpstina || '-'}</div>
            <div><b>GGAP:</b> ${p.GGAPStatus || '-'}</div>
            <div style="margin-top:6px;color:#666;">${p.ParcelaID}</div>
        </div>
    `;
}
    
async function loadParcele() {
    const parcele = (stammdaten.parcele || []).filter(p => p.KooperantID === CONFIG.ENTITY_ID);
    const list = document.getElementById('parceleList');
    const mapDiv = document.getElementById('parceleMap');
    
    if (parcele.length === 0) {
        list.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:40px;">Nema parcela</p>';
        mapDiv.style.display = 'none';
        return;
    }
    
    // Render list with loading placeholders
    list.innerHTML = parcele.map(p =>
        `<div id="parcel-card-${p.ParcelaID}" style="background:white;border-radius:var(--radius);padding:14px;margin-bottom:8px;box-shadow:0 1px 3px rgba(0,0,0,0.08);border-left:4px solid var(--primary);cursor:pointer;" onclick="focusParcel('${p.ParcelaID}')">
            <div style="display:flex;justify-content:space-between;margin-bottom:4px;">
                <strong>${p.KatBroj || p.ParcelaID}</strong>
                <span style="font-size:12px;color:var(--text-muted);">${p.ParcelaID}</span>
            </div>
            <div style="font-size:13px;color:var(--text-muted);margin-bottom:6px;">${p.Kultura || ''} | ${p.PovrsinaHa || '?'} ha | ${p.KatOpstina || ''}${p.GGAPStatus ? ' | GGAP: ' + p.GGAPStatus : ''}</div>
            <div id="parcel-meteo-${p.ParcelaID}" style="font-size:12px;color:var(--text-muted);">⏳ Meteo...</div>
        </div>`).join('');
    
    // Init map
    if (parcelMapInstance) { parcelMapInstance.remove(); parcelMapInstance = null; }
    parcelMapInstance = L.map(mapDiv).setView([43.28, 21.72], 13);
    L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}', {
        maxZoom: 22,
        attribution: 'Esri, Maxar, Earthstar Geographics'
    }).addTo(parcelMapInstance);
    
    // Load all parcel geo data + meteo
    const allBounds = [];
    window._parcelLayers = {};
    
    for (const p of parcele) {
        try {
            const resp = await fetch(CONFIG.API_URL + '?action=getParcelGeo&parcelaId=' + encodeURIComponent(p.ParcelaID));
            const json = await resp.json();
            
            if (json && json.success && json.parcel) {
                const geo = json.parcel;
                const lat = parseFloat(String(geo.Lat).replace(',', '.'));
                const lng = parseFloat(String(geo.Lng).replace(',', '.'));
                const popupHtml = buildKooperantParcelPopup(p);
                
                if (geo.PolygonGeoJSON) {
                    const geometry = JSON.parse(geo.PolygonGeoJSON);
                    const feature = {type: 'Feature', properties: p, geometry: geometry};
                    const layer = L.geoJSON(feature, { style: kooperantParcelStyle }).addTo(parcelMapInstance);
                    layer.eachLayer(l => {
                        l.bindTooltip(`${p.KatBroj || p.ParcelaID}`, { permanent: true, direction: 'center', className: 'parcel-label' });
                        l.bindPopup(popupHtml);
                        l.on('click', () => { highlightKooperantParcelLayer(l); });
                        const bounds = l.getBounds();
                        if (bounds.isValid()) allBounds.push(bounds);
                        window._parcelLayers[p.ParcelaID] = l;
                    });
                } else if (lat && lng && !isNaN(lat) && !isNaN(lng)) {
                    const marker = L.circleMarker([lat, lng], {
                        radius: 8, color: '#ffd60a', weight: 3, fillColor: '#ffd60a', fillOpacity: 0.85
                    }).addTo(parcelMapInstance);
                    marker.bindTooltip(`${p.KatBroj || p.ParcelaID}`, { permanent: true, direction: 'top', className: 'parcel-label' });
                    marker.bindPopup(popupHtml);
                    allBounds.push(L.latLngBounds([marker.getLatLng(), marker.getLatLng()]));
                    window._parcelLayers[p.ParcelaID] = marker;
                }
            }
        } catch (e) {}
        
        // Load meteo for this parcel (non-blocking per parcel)
        loadParcelMeteoInline(p.ParcelaID, p.Kultura || '');
    }

    if (allBounds.length > 0) {
        let combined = allBounds[0];
        for (let i = 1; i < allBounds.length; i++) {
            combined.extend(allBounds[i]);
        }
        parcelMapInstance.fitBounds(combined.pad(0.2));
    }
}


// ============================================================
// KOOPERANT: PARCELA METEO + RISK
// ============================================================
let meteoCache = {};

async function loadParcelMeteo(parcelaId, kultura) {
    const panel = document.getElementById('parceleMeteo');
    panel.style.display = 'block';
    panel.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:12px;">Učitavanje meteo podataka...</p>';
    
    // Check local cache
    if (meteoCache[parcelaId] && (Date.now() - meteoCache[parcelaId]._ts < 3600000)) {
        renderMeteoPanel(meteoCache[parcelaId]);
        return;
    }
    
    try {
        const url = CONFIG.API_URL + '?action=getParcelMeteo&parcelaId=' + encodeURIComponent(parcelaId);
        const resp = await fetch(url);
        const json = await resp.json();
        
        if (json && json.success) {
            json._ts = Date.now();
            meteoCache[parcelaId] = json;
            renderMeteoPanel(json);
        } else {
            panel.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:12px;">' + (json.error || 'Nema meteo podataka') + '</p>';
        }
    } catch (e) {
        panel.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:12px;">Greška pri učitavanju</p>';
    }
}

function renderMeteoPanel(data) {
    const panel = document.getElementById('parceleMeteo');
    const c = data.current || {};
    const risk = data.risk || {};
    const spray = data.sprayWindow || [];
    const daily = data.daily || data.ForecastDaily || [];
    const parcelaId = data.parcelaId || '';

    if (panel.dataset) panel.dataset.currentParcelaId = parcelaId;

    panel.innerHTML = `
        <div class="meteo-panel">
            <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;">
                <div style="font-size:14px;font-weight:700;color:var(--primary);">
                    ${data.katBroj || data.parcelaId} — ${data.kultura || ''}
                </div>
                <div style="font-size:10px;color:var(--text-muted);">
                    ${new Date(data.fetchedAt).toLocaleTimeString('sr', {hour:'2-digit',minute:'2-digit'})}
                </div>
            </div>
            
            <div class="meteo-current">
                <div>
                    <div class="meteo-temp">${Number(c.temperature || 0).toFixed(1)}°</div>
                </div>
                <div class="meteo-details">
                    <div>${weatherCodeText(c.weatherCode || 0)}</div>
                    <div>💧 Vlažnost: ${c.humidity || 0}%</div>
                    <div>💨 Vetar: ${Number(c.windSpeed || 0).toFixed(1)} km/h (udari: ${Number(c.windGusts || 0).toFixed(1)})</div>
                    ${Number(c.precipitation || 0) > 0 ? '<div>🌧️ Padavine: ' + Number(c.precipitation).toFixed(1) + ' mm</div>' : ''}
                </div>
            </div>
            
            ${renderRiskSection(risk)}
            ${renderSpraySection(spray, data.kultura)}
            ${renderForecast(daily, ['Ned', 'Pon', 'Uto', 'Sre', 'Čet', 'Pet', 'Sub'])}
            ${renderExpertInfo(parcelaId, c)}
        </div>
    `;
}
function renderRiskSection(risk) {
    if (!risk || !risk.items || risk.items.length === 0) {
        return '<div class="meteo-risk ok">✅ Nema rizika — uslovi su povoljni</div>';
    }
    
    return '<div style="margin-bottom:10px;">' +
        '<div style="font-size:12px;font-weight:600;color:var(--text-muted);margin-bottom:6px;">UPOZORENJA</div>' +
        risk.items.map(r =>
            `<div class="meteo-risk ${r.level}">
                <span style="font-size:18px;">${r.icon}</span>
                <span>${r.message}</span>
            </div>`
        ).join('') +
    '</div>';
}

function renderSpraySection(windows, kultura) {
    let html = '<div style="margin-bottom:10px;">';
    html += '<div style="font-size:12px;font-weight:600;color:var(--text-muted);margin-bottom:6px;">PROZOR ZA PRSKANJE</div>';
    
    if (!windows || windows.length === 0) {
        html += '<div class="spray-window" style="background:#fef3c7;border-color:#fcd34d;">⚠️ Nema pogodnog termina za prskanje u naredna 72h</div>';
    } else {
        windows.forEach((w, i) => {
            const start = new Date(w.start);
            const end = new Date(w.end);
            const startStr = start.toLocaleDateString('sr', {weekday:'short', day:'numeric', month:'short'}) + ' ' +
                           start.toLocaleTimeString('sr', {hour:'2-digit', minute:'2-digit'});
            const endStr = end.toLocaleTimeString('sr', {hour:'2-digit', minute:'2-digit'});
            
            html += `<div class="spray-window">
                <div class="spray-time">${i === 0 ? '✅ ' : ''}${startStr} — ${endStr} (${w.hours}h)</div>
                <div class="spray-details">Temp: ${w.avgTemp}°C | Vetar: ${w.avgWind} km/h | Vlažnost: ${w.avgHumidity}%</div>
            </div>`;
        });
    }
    
    html += '</div>';
    return html;
}

function renderForecast(daily, dayNames) {
    if (!daily || daily.length === 0) return '';

    const first3 = daily.slice(0, 3);

    let html = '<div style="font-size:12px;font-weight:600;color:var(--text-muted);margin-bottom:6px;">PROGNOZA — NAREDNA 3 DANA</div>';
    html += '<div class="meteo-forecast-3d">';

    first3.forEach((d, i) => {
        const date = new Date(d.date);
        const dayName = i === 0 ? 'Danas' : dayNames[date.getDay()];
        const icon = weatherCodeIcon(d.weatherCode);

        html += `
            <div class="meteo-day-3d">
                <div class="day-name">${dayName}</div>
                <div class="day-icon">${icon}</div>
                <div class="day-temp">${Math.round(d.tempMax)}°</div>
                <div class="day-temp-min">${Math.round(d.tempMin)}°</div>
                <div class="day-rain">${d.precipSum > 0 ? d.precipSum.toFixed(1) + ' mm' : '&nbsp;'}</div>
            </div>
        `;
    });

    html += '</div>';
    return html;
}

function weatherCodeText(code) {
    const codes = {
        0: 'Vedro', 1: 'Pretežno vedro', 2: 'Delimično oblačno', 3: 'Oblačno',
        45: 'Magla', 48: 'Magla sa mrazom',
        51: 'Slaba kiša', 53: 'Umerena kiša', 55: 'Jaka kiša',
        61: 'Slaba kiša', 63: 'Umerena kiša', 65: 'Jaka kiša',
        71: 'Slab sneg', 73: 'Umeren sneg', 75: 'Jak sneg',
        80: 'Pljuskovi', 81: 'Umereni pljuskovi', 82: 'Jaki pljuskovi',
        95: 'Grmljavina', 96: 'Grmljavina sa gradom', 99: 'Jaka grmljavina'
    };
    return codes[code] || 'Nepoznato';
}

function weatherCodeIcon(code) {
    if (code === 0) return '☀️';
    if (code <= 3) return '⛅';
    if (code <= 48) return '🌫️';
    if (code <= 65) return '🌧️';
    if (code <= 75) return '❄️';
    if (code <= 82) return '🌦️';
    if (code >= 95) return '⛈️';
    return '🌤️';
}

async function loadParcelMeteoInline(parcelaId, kultura) {
    const el = document.getElementById('parcel-meteo-' + parcelaId);
    if (!el) return;
    
    // Check local cache
    if (meteoCache[parcelaId] && (Date.now() - meteoCache[parcelaId]._ts < 3600000)) {
        el.innerHTML = renderMeteoInline(meteoCache[parcelaId]);
        return;
    }
    
    try {
        const url = CONFIG.API_URL + '?action=getParcelMeteo&parcelaId=' + encodeURIComponent(parcelaId);
        const resp = await fetch(url);
        const json = await resp.json();
        
        if (json && json.success) {
            json._ts = Date.now();
            meteoCache[parcelaId] = json;
            el.innerHTML = renderMeteoInline(json);
        } else {
            el.innerHTML = '<span style="color:var(--text-muted);">Nema meteo podataka</span>';
        }
    } catch (e) {
        el.innerHTML = '<span style="color:var(--text-muted);">—</span>';
    }
}

function renderMeteoInline(data) {
    const c = data.current || {};
    const risk = data.risk || {};
    const spray = data.sprayWindow || data.SprayWindows || [];
    const daily = data.daily || data.ForecastDaily || [];

    const temp = Number(c.temperature || c.Temp || 0).toFixed(1);
    const hum = Number(c.humidity || c.Humidity || 0).toFixed(0);
    const wind = Number(c.windSpeed || c.Wind || 0).toFixed(0);
    const icon = weatherCodeIcon(c.weatherCode || c.WeatherCode || 0);

    let riskHtml = '<span class="parcel-chip ok">✅ Bez rizika</span>';
    const riskItems = risk ? (risk.items || risk.RiskItems || []) : [];
    if (riskItems.length > 0) {
        const first = riskItems[0];
        const cls = first.level === 'danger' ? 'danger' : 'warn';
        riskHtml = `<span class="parcel-chip ${cls}">${first.icon} ${first.message}</span>`;
    }

    let sprayHtml = '<span class="parcel-chip warn">⚠️ Nema termina za prskanje</span>';
    if (spray.length > 0) {
        const w = spray[0];
        const start = new Date(w.start);
        const end = new Date(w.end);
        const dayNames = ['Ned','Pon','Uto','Sre','Čet','Pet','Sub'];
        const dayStr = dayNames[start.getDay()];
        const startTime = start.toLocaleTimeString('sr', { hour:'2-digit', minute:'2-digit' });
        const endTime = end.toLocaleTimeString('sr', { hour:'2-digit', minute:'2-digit' });
        sprayHtml = `<span class="parcel-chip ok">🎯 ${dayStr} ${startTime}-${endTime} (${w.hours}h)</span>`;
    }

    let forecastHtml = '';
    if (daily.length > 0) {
        const dayNames = ['Ned','Pon','Uto','Sre','Čet','Pet','Sub'];
        forecastHtml = `
            <div class="parcel-forecast-inline">
                ${daily.slice(0, 3).map((d, i) => {
                    const dt = new Date(d.date);
                    const name = i === 0 ? 'Danas' : dayNames[dt.getDay()];
                    return `
                        <span class="parcel-forecast-day">
                            <strong>${name}</strong>
                            ${weatherCodeIcon(d.weatherCode || 0)}
                            ${Math.round(Number(d.tempMax || 0))}°/${Math.round(Number(d.tempMin || 0))}°
                            ${Number(d.precipSum || 0) > 0 ? '💧' + Number(d.precipSum).toFixed(1) : ''}
                        </span>
                    `;
                }).join('')}
            </div>
        `;
    }

    return `
        <div class="parcel-meteo-compact">
            <div class="parcel-meteo-topline">
                <span class="parcel-chip temp">${temp}°C</span>
                <span class="parcel-chip">${icon}</span>
                <span class="parcel-chip">💧 ${hum}%</span>
                <span class="parcel-chip">💨 ${wind} km/h</span>
                ${riskHtml}
            </div>
            <div class="parcel-meteo-midline">
                ${sprayHtml}
            </div>
            <div class="parcel-meteo-bottomline">
                ${forecastHtml}
            </div>
        </div>
    `;
}

function renderExpertInfo(parcelId, current) {
    const soilMoist = current.soilMoist_0_1 ?? current.SoilMoist_0_1cm ?? null;
    const soilTemp = current.soilTemp_0 ?? current.SoilTemp_0cm ?? null;
    const et0 = current.et0 ?? current.ET0 ?? null;
    const uv = current.uvIndex ?? current.UVIndex ?? null;
    const solar = current.solarRadiation ?? current.SolarRadiation ?? null;

    const hasExpert =
        soilMoist !== null ||
        soilTemp !== null ||
        et0 !== null ||
        uv !== null ||
        solar !== null;

    if (!hasExpert) return '';

    const isOpen = !!parcelExpertOpen[parcelId];

    let html = `
        <button class="parcel-expert-toggle" onclick="toggleParcelExpert('${parcelId}')">
            <div>
                <div class="parcel-expert-title">
                    <span>🧪 Expert info</span>
                </div>
                <div class="parcel-expert-sub">Zemljište, ET₀, UV i dodatni agro podaci</div>
            </div>
            <span class="parcel-expert-chevron ${isOpen ? 'open' : ''}">⌄</span>
        </button>
    `;

    if (!isOpen) return html;

    html += `
        <div class="parcel-expert-panel">
            <div class="parcel-expert-grid">
                ${soilTemp !== null ? `
                    <div class="parcel-expert-item">
                        <div class="parcel-expert-k">🌡️ Temperatura zemljišta</div>
                        <div class="parcel-expert-v">${Number(soilTemp).toFixed(1)}°C</div>
                    </div>
                ` : ''}

                ${soilMoist !== null ? `
                    <div class="parcel-expert-item">
                        <div class="parcel-expert-k">🌱 Vlažnost zemljišta</div>
                        <div class="parcel-expert-v">${(Number(soilMoist) * 100).toFixed(0)}%</div>
                    </div>
                ` : ''}

                ${et0 !== null ? `
                    <div class="parcel-expert-item">
                        <div class="parcel-expert-k">💦 ET₀</div>
                        <div class="parcel-expert-v">${Number(et0).toFixed(1)} mm</div>
                    </div>
                ` : ''}

                ${uv !== null && Number(uv) > 0 ? `
                    <div class="parcel-expert-item">
                        <div class="parcel-expert-k">☀️ UV indeks</div>
                        <div class="parcel-expert-v">${Number(uv).toFixed(1)}</div>
                    </div>
                ` : ''}

                ${solar !== null && Number(solar) > 0 ? `
                    <div class="parcel-expert-item">
                        <div class="parcel-expert-k">🔆 Solarno zračenje</div>
                        <div class="parcel-expert-v">${Number(solar).toFixed(0)} W/m²</div>
                    </div>
                ` : ''}
            </div>
        </div>
    `;

    return html;
}

function toggleParcelExpert(parcelId) {
    parcelExpertOpen[parcelId] = !parcelExpertOpen[parcelId];

    const panel = document.getElementById('parceleMeteo');
    if (panel && panel.dataset && panel.dataset.currentParcelaId === parcelId) {
        const cached = meteoCache[parcelId];
        if (cached) renderMeteoPanel(cached);
    }

    const parcela = (stammdaten.parcele || []).find(p => p.ParcelaID === parcelId);
    if (parcela) {
        loadParcelMeteoInline(parcelId, parcela.Kultura || '');
    }
}

function focusParcel(parcelaID) {
    if (!parcelMapInstance || !window._parcelLayers || !window._parcelLayers[parcelaID]) return;

    const layer = window._parcelLayers[parcelaID];

    if (layer.getBounds) {
        parcelMapInstance.fitBounds(layer.getBounds().pad(0.3));
        highlightKooperantParcelLayer(layer);
    } else if (layer.getLatLng) {
        parcelMapInstance.setView(layer.getLatLng(), 17);
        resetKooperantParcelHighlight();
    }

    if (layer.openPopup) layer.openPopup();

    document.getElementById('parceleMap').scrollIntoView({ behavior: 'smooth' });
}
