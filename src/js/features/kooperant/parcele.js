// ============================================================
// KOOPERANT: PARCELE
// ============================================================
let parcelMapInstance = null;
let parcelLayers = {};
let _parceleLoaded = false;

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
                ${escapeHtml(p.KatBroj || p.ParcelaID)}
            </div>
            <div><b>Kultura:</b> ${escapeHtml(p.Kultura || '-')}</div>
            <div><b>Površina:</b> ${escapeHtml(p.PovrsinaHa || '?')} ha</div>
            <div><b>KO:</b> ${escapeHtml(p.KatOpstina || '-')}</div>
            <div><b>GGAP:</b> ${escapeHtml(p.GGAPStatus || '-')}</div>
            <div style="margin-top:6px;color:#666;">${escapeHtml(p.ParcelaID)}</div>
        </div>
    `;
}

async function loadParcele() {
    // Ako je već učitano — ne radi ništa
    if (_parceleLoaded && parcelMapInstance) {
        return;
    }

    await safeAsync(async () => {
        const parcele = (stammdaten.parcele || []).filter(p => p.KooperantID === CONFIG.ENTITY_ID);
        const list = document.getElementById('parceleList');
        const mapDiv = document.getElementById('parceleMap');

        if (!list || !mapDiv) return;

        // Prepopulate meteoCache iz stammdaten — ODMAH, bez API poziva
        (stammdaten.meteoLatest || []).forEach(m => {
            const pid = String(m.ParcelaID || '').trim();
            if (!pid) return;
            if (window.meteoCache[pid] && (Date.now() - window.meteoCache[pid]._ts < METEO_CACHE_TTL)) return;

            try {
                const riskItems = typeof m.RiskItems === 'string' ? JSON.parse(m.RiskItems || '[]') : (m.RiskItems || []);
                const sprayWindows = typeof m.SprayWindows === 'string' ? JSON.parse(m.SprayWindows || '[]') : (m.SprayWindows || []);
                const forecastDaily = typeof m.ForecastDaily === 'string' ? JSON.parse(m.ForecastDaily || '[]') : (m.ForecastDaily || []);

                window.meteoCache[pid] = {
                    success: true,
                    parcelaId: pid,
                    kultura: m.Kultura || '',
                    fetchedAt: m.LastFetch || '',
                    _ts: Date.now(),
                    current: {
                        temperature: Number(m.Temp) || 0,
                        humidity: Number(m.Humidity) || 0,
                        windSpeed: Number(m.Wind) || 0,
                        windGusts: Number(m.WindGusts) || 0,
                        precipitation: Number(m.Precip) || 0,
                        weatherCode: Number(m.WeatherCode) || 0,
                        dewPoint: Number(m.DewPoint) || 0,
                        cloudCover: Number(m.CloudCover) || 0,
                        uvIndex: Number(m.UVIndex) || 0,
                        solarRadiation: Number(m.SolarRadiation) || 0,
                        soilMoist_0_1: Number(m.SoilMoist_0_1cm) || 0,
                        soilTemp_0: Number(m.SoilTemp_0cm) || 0,
                        et0: Number(m.ET0) || 0
                    },
                    risk: {
                        level: m.RiskLevel || 'ok',
                        items: riskItems
                    },
                    sprayWindow: sprayWindows,
                    daily: forecastDaily
                };
            } catch (e) {
                console.error('meteoLatest parse error for ' + pid, e);
            }
        });
        
        if (parcele.length === 0) {
            list.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:40px;">Nema parcela</p>';
            mapDiv.style.display = 'none';
            return;
        }

        mapDiv.style.display = '';

        list.innerHTML = parcele.map(p =>
            `<div id="parcel-card-${p.ParcelaID}" style="background:white;border-radius:var(--radius);padding:14px;margin-bottom:8px;box-shadow:0 1px 3px rgba(0,0,0,0.08);border-left:4px solid var(--primary);cursor:pointer;" onclick="focusParcel('${String(p.ParcelaID).replace(/'/g, "\\'")}')">
                <div style="display:flex;justify-content:space-between;margin-bottom:4px;">
                    <strong>${escapeHtml(p.KatBroj || p.ParcelaID)}</strong>
                    <span style="font-size:12px;color:var(--text-muted);">${escapeHtml(p.ParcelaID)}</span>
                </div>
                <div style="font-size:13px;color:var(--text-muted);margin-bottom:6px;">
                    ${escapeHtml(p.Kultura || '')} |
                    ${escapeHtml(p.PovrsinaHa || '?')} ha |
                    ${escapeHtml(p.KatOpstina || '')}
                    ${p.GGAPStatus ? ' | GGAP: ' + escapeHtml(p.GGAPStatus) : ''}
                </div>
                <div id="parcel-meteo-${p.ParcelaID}" style="font-size:12px;color:var(--text-muted);">⏳ Meteo...</div>
            </div>`
        ).join('');

        if (parcelMapInstance) {
            parcelMapInstance.remove();
            parcelMapInstance = null;
        }

        parcelMapInstance = L.map(mapDiv).setView([43.28, 21.72], 13);
        L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}', {
            maxZoom: 22,
            attribution: 'Esri, Maxar, Earthstar Geographics'
        }).addTo(parcelMapInstance);

        const allBounds = [];
        parcelLayers = {};

        const geoResults = await Promise.all(parcele.map(async (p) => {
            const json = await safeAsync(async () => {
                return await apiFetch('action=getParcelGeo&parcelaId=' + encodeURIComponent(p.ParcelaID));
            });
            return { parcela: p, json };
        }));

        geoResults.forEach(({ parcela: p, json }) => {
            if (!(json && json.success && json.parcel)) return;

            const geo = json.parcel;
            const lat = parseFloat(String(geo.Lat).replace(',', '.'));
            const lng = parseFloat(String(geo.Lng).replace(',', '.'));
            const popupHtml = buildKooperantParcelPopup(p);

            // Preskoči parcele bez validnih koordinata
            if (!lat || !lng || isNaN(lat) || isNaN(lng) || (lat === 0 && lng === 0)) return;

            if (geo.PolygonGeoJSON) {
                try {
                    const geometry = JSON.parse(geo.PolygonGeoJSON);
                    const feature = { type: 'Feature', properties: p, geometry: geometry };
                    const layer = L.geoJSON(feature, { style: kooperantParcelStyle }).addTo(parcelMapInstance);

                    layer.eachLayer(l => {
                        l.bindTooltip(`${p.KatBroj || p.ParcelaID}`, { permanent: true, direction: 'center', className: 'parcel-label' });
                        l.bindPopup(popupHtml);
                        l.on('click', () => { highlightKooperantParcelLayer(l); });

                        const bounds = l.getBounds();
                        if (bounds && bounds.isValid()) allBounds.push(bounds);

                        parcelLayers[p.ParcelaID] = l;
                    });
                } catch (e) {
                    console.error('Invalid polygon geojson for', p.ParcelaID, e);
                }
            } else if (!isNaN(lat) && !isNaN(lng)) {
                const marker = L.circleMarker([lat, lng], {
                    radius: 8,
                    color: '#ffd60a',
                    weight: 3,
                    fillColor: '#ffd60a',
                    fillOpacity: 0.85
                }).addTo(parcelMapInstance);

                marker.bindTooltip(`${p.KatBroj || p.ParcelaID}`, { permanent: true, direction: 'top', className: 'parcel-label' });
                marker.bindPopup(popupHtml);

                allBounds.push(L.latLngBounds([marker.getLatLng(), marker.getLatLng()]));
                parcelLayers[p.ParcelaID] = marker;
            }
        });

        // Meteo fallback — samo za parcele koje NEMAJU cached podatke
        parcele.forEach(p => {
            if (!window.meteoCache[p.ParcelaID]) {
                loadParcelMeteoInline(p.ParcelaID, p.Kultura || '');
            }
        });

        if (allBounds.length > 0) {
            let combined = allBounds[0];
            for (let i = 1; i < allBounds.length; i++) {
                combined.extend(allBounds[i]);
            }
            parcelMapInstance.fitBounds(combined.pad(0.2));
        }

        _parceleLoaded = true;
        populateParceleKulturaFilter();
    }, 'Greška pri učitavanju parcela');
}

// Invalidate kad se stammdaten refreshuju
function invalidateParceleCache() {
    _parceleLoaded = false;
}
// ============================================================
// KOOPERANT: PARCELA METEO + RISK
// ============================================================
window.meteoCache = window.meteoCache || {};
const METEO_CACHE_TTL = 6 * 60 * 60 * 1000;

async function loadParcelMeteo(parcelaId, kultura) {
    const panel = document.getElementById('parceleMeteo');
    if (!panel) return;

    panel.style.display = 'block';
    panel.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:12px;">Učitavanje meteo podataka...</p>';

    if (window.meteoCache[parcelaId] && (Date.now() - window.meteoCache[parcelaId]._ts < METEO_CACHE_TTL)) {
        renderMeteoPanel(window.meteoCache[parcelaId]);
        return;
    }
    
    const json = await safeAsync(async () => {
        return await apiFetch('action=getParcelMeteo&parcelaId=' + encodeURIComponent(parcelaId));
    }, 'Greška pri učitavanju meteo podataka');

    if (!json) {
        panel.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:12px;">Greška pri učitavanju</p>';
        return;
    }

    if (json.success) {
        json._ts = Date.now();
        window.meteoCache[parcelaId] = json;
        renderMeteoPanel(json);
    } else {
        panel.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:12px;">' + escapeHtml(json.error || 'Nema meteo podataka') + '</p>';
    }
}

function renderMeteoPanel(data) {
    const panel = document.getElementById('parceleMeteo');
    if (!panel) return;

    const c = data.current || {};
    const risk = data.risk || {};
    const spray = data.sprayWindow || [];
    const daily = data.daily || data.ForecastDaily || [];
    const parcelaId = data.parcelaId || '';

    if (panel.dataset) panel.dataset.currentParcelaId = parcelaId;

    let fetchedTime = '';
    try {
        fetchedTime = data.fetchedAt
            ? new Date(data.fetchedAt).toLocaleTimeString('sr', { hour: '2-digit', minute: '2-digit' })
            : '';
    } catch (_) {
        fetchedTime = '';
    }

    panel.innerHTML = `
        <div class="meteo-panel">
            <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;">
                <div style="font-size:14px;font-weight:700;color:var(--primary);">
                    ${escapeHtml(data.katBroj || data.parcelaId || '')} — ${escapeHtml(data.kultura || '')}
                </div>
                <div style="font-size:10px;color:var(--text-muted);">
                    ${escapeHtml(fetchedTime)}
                </div>
            </div>
            
            <div class="meteo-current">
                <div>
                    <div class="meteo-temp">${Number(c.temperature || 0).toFixed(1)}°</div>
                </div>
                <div class="meteo-details">
                    <div>${escapeHtml(weatherCodeText(c.weatherCode || 0))}</div>
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
                <span>${escapeHtml(r.message)}</span>
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

    // Ako imamo cached podatke mlađe od 6h — prikaži iz keša, ne zovi API
    if (window.meteoCache[parcelaId] && (Date.now() - window.meteoCache[parcelaId]._ts < METEO_CACHE_TTL)) {
        el.innerHTML = renderMeteoInline(window.meteoCache[parcelaId], parcelaId);
        return;
    }

    const json = await safeAsync(async () => {
        return await apiFetch('action=getParcelMeteo&parcelaId=' + encodeURIComponent(parcelaId));
    });

    if (!json) {
        el.innerHTML = '<span style="color:var(--text-muted);">—</span>';
        return;
    }

    if (json && json.success) {
        json._ts = Date.now();
        window.meteoCache[parcelaId] = json;
        el.innerHTML = renderMeteoInline(json, parcelaId);
    } else {
        el.innerHTML = '<span style="color:var(--text-muted);">Nema meteo podataka</span>';
    }
}

function renderMeteoInline(data, parcelaId) {
    const c = data.current || {};
    const pid = parcelaId || data.parcelaId || '';
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
        riskHtml = `<span class="parcel-chip ${cls}">${first.icon} ${escapeHtml(first.message)}</span>`;
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

    // Expert info — dodato
    const expertHtml = renderExpertInfo(pid, c);
    
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
            ${expertHtml}
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

    return `
        <div class="parcel-expert-wrapper" id="expert-wrapper-${parcelId}">
            <button class="parcel-expert-toggle" onclick="event.stopPropagation(); toggleExpertPanel('${String(parcelId).replace(/'/g, "\\'")}')">
                <div>
                    <div class="parcel-expert-title"><span>🧪 Expert info</span></div>
                    <div class="parcel-expert-sub">Zemljište, ET₀, UV i dodatni agro podaci</div>
                </div>
                <span class="parcel-expert-chevron" id="expert-chevron-${parcelId}">⌄</span>
            </button>
            <div class="parcel-expert-panel" id="expert-panel-${parcelId}" style="display:none;">
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
        </div>
    `;
}

let _activeExpertPanel = null;

function toggleExpertPanel(parcelId) {
    const panel = document.getElementById('expert-panel-' + parcelId);
    const chevron = document.getElementById('expert-chevron-' + parcelId);
    if (!panel) return;

    // Ako kliknemo na isti koji je otvoren — zatvori
    if (_activeExpertPanel === parcelId) {
        panel.style.display = 'none';
        if (chevron) chevron.classList.remove('open');
        _activeExpertPanel = null;
        return;
    }

    // Zatvori prethodni ako postoji
    closeActiveExpertPanel();

    // Otvori novi
    panel.style.display = 'block';
    if (chevron) chevron.classList.add('open');
    _activeExpertPanel = parcelId;
}

function closeActiveExpertPanel() {
    if (!_activeExpertPanel) return;
    const panel = document.getElementById('expert-panel-' + _activeExpertPanel);
    const chevron = document.getElementById('expert-chevron-' + _activeExpertPanel);
    if (panel) panel.style.display = 'none';
    if (chevron) chevron.classList.remove('open');
    _activeExpertPanel = null;
}

// Klik bilo gde van expert panela — zatvori
document.addEventListener('click', function(e) {
    if (!_activeExpertPanel) return;
    const wrapper = document.getElementById('expert-wrapper-' + _activeExpertPanel);
    if (wrapper && !wrapper.contains(e.target)) {
        closeActiveExpertPanel();
    }
});

function focusParcel(parcelaID) {
    const mapDiv = document.getElementById('parceleMap');
    if (!parcelMapInstance || !parcelLayers[parcelaID]) return;

    const layer = parcelLayers[parcelaID];

    if (layer.getBounds) {
        parcelMapInstance.fitBounds(layer.getBounds().pad(0.3));
        highlightKooperantParcelLayer(layer);
    } else if (layer.getLatLng) {
        parcelMapInstance.setView(layer.getLatLng(), 17);
        resetKooperantParcelHighlight();
    }

    if (layer.openPopup) layer.openPopup();
    if (mapDiv) mapDiv.scrollIntoView({ behavior: 'smooth' });
}


function showParceleSection(name, btn) {
    document.querySelectorAll('.parcele-section').forEach(el => el.classList.remove('active'));
    document.querySelectorAll('.parcele-subnav-btn').forEach(el => el.classList.remove('active'));

    const section = document.getElementById('parcele-section-' + name);
    if (section) section.classList.add('active');

    if (btn && btn.classList) {
        btn.classList.add('active');
    }

    const toggleBtn = document.querySelector('.parcele-primary-btn');
    if (toggleBtn) {
        toggleBtn.textContent = name === 'mapa' ? 'Prikaži listu' : 'Prikaži mapu';
    }
}

function toggleParceleView() {
    const mapaActive = document.getElementById('parcele-section-mapa')?.classList.contains('active');
    const buttons = document.querySelectorAll('.parcele-subnav-btn');

    if (mapaActive) {
        showParceleSection('lista', buttons[1] || null);
    } else {
        showParceleSection('mapa', buttons[0] || null);
    }
}

function populateParceleKulturaFilter() {
    const sel = document.getElementById('parceleKulturaFilter');
    if (!sel) return;
    if (sel.options.length > 1) return;

    const values = Array.from(new Set(
        (stammdaten.parcele || [])
            .filter(p => p.KooperantID === CONFIG.ENTITY_ID)
            .map(p => String(p.Kultura || '').trim())
            .filter(Boolean)
    )).sort();

    values.forEach(v => {
        const o = document.createElement('option');
        o.value = v;
        o.textContent = v;
        sel.appendChild(o);
    });
}

function applyParceleFilters() {
    const search = String(document.getElementById('parceleSearch')?.value || '').toLowerCase().trim();
    const kultura = String(document.getElementById('parceleKulturaFilter')?.value || '').trim();
    const list = document.getElementById('parceleList');
    if (!list) return;

    Array.from(list.children).forEach(item => {
        const text = (item.innerText || '').toLowerCase();
        const matchSearch = !search || text.includes(search);
        const matchKultura = !kultura || text.includes(kultura.toLowerCase());
        item.style.display = (matchSearch && matchKultura) ? '' : 'none';
    });
}
