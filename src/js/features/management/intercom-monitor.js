/**
 * intercom-monitor.js
 * AgriX PWA — Management intercom monitor module (Phase 1)
 *
 * Design contract:
 *   - Management sees list of stations within working hours
 *   - "Slušaj" on a station card writes /intercom/requests/{otkupacID}/status="requested"
 *   - Waits for status="active" (Otkupac mic confirmed live) then starts chunk listener
 *   - Handles status="error" (mic permission denied on Otkupac side)
 *   - Only one station listened to at a time
 *   - "Stop" writes status="ended" → Otkupac mic stops
 *
 * Session list source:
 *   - Working-hours filter is CALLER responsibility (stammdaten.stanice radnoVreme)
 *   - This module receives a pre-filtered array of { entityID, name, stationID }
 *     via initIntercomMonitor(stations) and owns only Firebase + audio logic
 *
 * Mount point: #intercom-monitor-section inside tab-mgmt-dispecer
 *
 * Dependencies (window globals):
 *   CONFIG     — FIREBASE_* keys
 *   showToast, byId, showEl, hideEl, setText, escapeHtml
 *   firebase   — Firebase compat UMD
 *
 * NOT in scope (Phase 1):
 *   - fromMgmt back-channel
 *   - Custom token auth
 */

'use strict';

(function () {

// ─── Module state ─────────────────────────────────────────────────────────────

const _m = {
  db:           null,
  audioCtx:     null,
  stations:     [],      // [{ entityID, name, stationID }] — passed in at init
  requestUnsub: null,    // listener on active request node
  listeningTo:  null,    // entityID currently being listened to
  listenUnsub:  null,    // chunk listener detach fn
  lastSeq:      -1,
  playQueue:    [],
  playing:      false,
  reqStatus:    'idle'   // 'idle'|'requesting'|'active'|'error'
};

// ─── Firebase init + anon auth ────────────────────────────────────────────────

function _initFirebase() {
  if (_f.db) return;

  if (typeof firebase === 'undefined') {
    throw new Error('Firebase global nije učitan');
  }
  if (typeof firebase.initializeApp !== 'function') {
    throw new Error('firebase-app-compat.js nije učitan');
  }

  // Forsiraj WebSocket — sprečava long-polling i frame fallback (čistiji CSP)
  if (firebase.database?.INTERNAL?.forceWebSockets) {
    firebase.database.INTERNAL.forceWebSockets();
  }
  
  const existing = firebase.apps.find(a => a.name === 'agrix-intercom');
  const app = existing || firebase.initializeApp({
    apiKey:      CONFIG.FIREBASE_API_KEY,
    projectId:   CONFIG.FIREBASE_PROJECT_ID,
    appId:       CONFIG.FIREBASE_APP_ID,
    databaseURL: CONFIG.FIREBASE_RTDB_URL
  }, 'agrix-intercom');

  if (typeof app.database !== 'function') {
    throw new Error('firebase-database-compat.js nije učitan');
  }
  if (typeof app.auth !== 'function') {
    throw new Error('firebase-auth-compat.js nije učitan');
  }

  _m.db = app.database();

  if (!_m.db) {
    throw new Error('app.database() vratio null/undefined');
  }
}

async function _ensureAuth() {
  const auth = firebase.app('agrix-intercom').auth();
  if (auth.currentUser) return;
  await auth.signInAnonymously();
}

// ─── Bootstrap ────────────────────────────────────────────────────────────────
// stations: [{ entityID, name, stationID }]
// Called from management bootstrap after stammdaten is loaded.

async function initIntercomMonitor(stations) {
  try {
    _initFirebase();
    await _ensureAuth();
    _m.stations = stations || [];
    renderMonitorSection();
  } catch (err) {
    console.error('[intercom-monitor] init failed:', err);
    showToast('Interkom monitor — greška inicijalizacije', 'error');
  }
}

// Called when stammdaten refreshes or working-hours filter changes
function updateIntercomStations(stations) {
  _m.stations = stations || [];
  renderMonitorSection();
}

// ─── Request flow ─────────────────────────────────────────────────────────────

async function requestListen(entityID) {
  if (_m.reqStatus !== 'idle') return;

  // AudioContext must be created from user gesture (this click handler is it)
  if (!_m.audioCtx) {
    _m.audioCtx = new (window.AudioContext || window.webkitAudioContext)();
  }
  if (_m.audioCtx.state === 'suspended') {
    await _m.audioCtx.resume();
  }

  // Stop any existing session first
  if (_m.listeningTo) await _stopListen(false);

  _m.listeningTo = entityID;
  _m.reqStatus   = 'requesting';
  _renderCard(entityID, 'requesting');

  try {
    // Write request — Otkupac listener will pick this up automatically
    await _m.db.ref(`intercom/requests/${entityID}`).set({
      status:      'requested',
      requestedAt: firebase.database.ServerValue.TIMESTAMP,
      requestedBy: CONFIG.ENTITY_NAME || 'Management'
    });

    // Listen for status changes on this request
    _startRequestStatusListener(entityID);

  } catch (err) {
    console.error('[intercom-monitor] requestListen error:', err);
    showToast('Greška slanja zahteva', 'error');
    _resetState();
  }
}

function _startRequestStatusListener(entityID) {
  if (_m.requestUnsub) {
    _m.requestUnsub();
    _m.requestUnsub = null;
  }

  const statusRef = _m.db.ref(`intercom/requests/${entityID}/status`);
  const handler = statusRef.on('value', snapshot => {
    const status = snapshot.val();
    if (!status) return;

    switch (status) {
      case 'active':
        _m.reqStatus = 'active';
        _renderCard(entityID, 'active');
        _renderListenBanner(entityID);
        _startChunkListener(entityID);
        break;

      case 'ended':
        _onRemoteEnd();
        break;

      case 'error':
        showToast(`${_stationName(entityID)} — mikrofon nije dostupan`, 'error');
        _resetState();
        break;
    }
  });

  _m.requestUnsub = () => statusRef.off('value', handler);
}

// ─── Chunk listener + playback ────────────────────────────────────────────────

function _startChunkListener(entityID) {
  if (_m.listenUnsub) {
    _m.listenUnsub();
    _m.listenUnsub = null;
  }

  _m.lastSeq   = -1;
  _m.playQueue = [];
  _m.playing   = false;

  const chunkRef = _m.db.ref(`intercom/sessions/${entityID}/toMgmt`);
  const handler = chunkRef.on('value', snapshot => {
    const data = snapshot.val();
    if (!data?.chunk) return;
    if (data.seq <= _m.lastSeq) return;   // deduplicate
    _m.lastSeq = data.seq;
    _decodeAndEnqueue(data.chunk);
  });

  _m.listenUnsub = () => chunkRef.off('value', handler);
}

function _decodeAndEnqueue(base64) {
  if (!_m.audioCtx) return;

  let binary;
  try { binary = atob(base64); }
  catch { return; }

  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);

  _m.audioCtx.decodeAudioData(bytes.buffer.slice(0),
    decoded => {
      _m.playQueue.push(decoded);
      if (!_m.playing) _drainQueue();
    },
    err => console.warn('[intercom-monitor] decodeAudioData:', err)
  );
}

function _drainQueue() {
  if (_m.playQueue.length === 0) { _m.playing = false; return; }
  _m.playing = true;
  const buf = _m.playQueue.shift();
  const src = _m.audioCtx.createBufferSource();
  src.buffer = buf;
  src.connect(_m.audioCtx.destination);
  src.onended = _drainQueue;
  src.start();
}

// ─── Stop listen ──────────────────────────────────────────────────────────────

async function _stopListen(writeEnded = true) {
  const entityID = _m.listeningTo;

  if (_m.listenUnsub)  { _m.listenUnsub();  _m.listenUnsub  = null; }
  if (_m.requestUnsub) { _m.requestUnsub(); _m.requestUnsub = null; }

  if (writeEnded && entityID) {
    try {
      await _m.db.ref(`intercom/requests/${entityID}/status`).set('ended');
    } catch (err) {
      console.warn('[intercom-monitor] _stopListen write error:', err);
    }
  }

  _resetState();
}

function _onRemoteEnd() {
  // Otkupac side ended the session (app closed, disconnect, etc.)
  const name = _stationName(_m.listeningTo);
  if (_m.reqStatus === 'active') showToast(`${name} — sesija završena`, 'info');
  if (_m.listenUnsub)  { _m.listenUnsub();  _m.listenUnsub  = null; }
  if (_m.requestUnsub) { _m.requestUnsub(); _m.requestUnsub = null; }
  _resetState();
}

function _resetState() {
  const prev       = _m.listeningTo;
  _m.listeningTo   = null;
  _m.reqStatus     = 'idle';
  _m.playQueue     = [];
  _m.playing       = false;
  _m.lastSeq       = -1;
  if (prev) _renderCard(prev, 'idle');
  _renderListenBanner(null);
}

// ─── UI rendering ─────────────────────────────────────────────────────────────

function renderMonitorSection() {
  const list = byId('intercom-session-list');
  if (!list) return;

  if (_m.stations.length === 0) {
    list.innerHTML = `<p class="intercom-empty">Nema stanica u radnom vremenu</p>`;
    return;
  }

  list.innerHTML = _m.stations.map(s => _cardHtml(s, _cardState(s.entityID))).join('');
}

function _cardState(entityID) {
  if (_m.listeningTo !== entityID) return 'idle';
  return _m.reqStatus;   // 'requesting' | 'active'
}

function _renderCard(entityID, state) {
  const station = _m.stations.find(s => s.entityID === entityID);
  if (!station) return;
  const existing = byId(`intercom-card-${entityID}`);
  if (!existing) { renderMonitorSection(); return; }
  existing.outerHTML = _cardHtml(station, state);
}

function _cardHtml(station, state) {
  const id  = escapeHtml(station.entityID);
  const nm  = escapeHtml(station.name);
  const st  = escapeHtml(station.stationID);

  let actionHtml;
  switch (state) {
    case 'requesting':
      actionHtml = `<span class="intercom-waiting">⏳ Čekanje...</span>`;
      break;
    case 'active':
      actionHtml = `
        <span class="intercom-live-label">🎧 Aktivno</span>
        <button class="btn-sm btn-danger"
                data-action="intercom-stop"
                data-entity-id="${id}">■ Stop</button>`;
      break;
    default:
      actionHtml = `
        <button class="btn-sm btn-primary"
                data-action="intercom-listen"
                data-entity-id="${id}">🎧 Slušaj</button>`;
  }

  return `
    <div id="intercom-card-${id}"
         class="intercom-session-card ${state === 'active' ? 'intercom-card--active' : ''}">
      <div class="intercom-card-info">
        <span class="intercom-live-dot ${state === 'active' ? 'intercom-dot--live' : ''}"></span>
        <strong>${nm}</strong>
        <span class="intercom-station">${st}</span>
      </div>
      <div class="intercom-card-actions">${actionHtml}</div>
    </div>`;
}

function _renderListenBanner(entityID) {
  const banner   = byId('intercom-active-listen-banner');
  const nameSpan = byId('intercom-listen-name');
  if (!banner) return;

  if (entityID) {
    setText(nameSpan, _stationName(entityID));
    showEl(banner);
  } else {
    hideEl(banner);
  }
}

function _stationName(entityID) {
  const s = _m.stations.find(s => s.entityID === entityID);
  return s ? s.name : entityID;
}

// ─── data-action handler (routed from app.js handleAppShellClick) ─────────────

function handleIntercomMonitorAction(action, dataset) {
  switch (action) {
    case 'intercom-listen':
      requestListen(dataset.entityId);
      break;
    case 'intercom-stop':
      _stopListen(true);
      break;
  }
}

// ─── Cleanup ──────────────────────────────────────────────────────────────────

function destroyIntercomMonitor() {
  _stopListen(false);
  _m.audioCtx?.close();
  _m.db     = null;
  _m.audioCtx = null;
}

// ─── Public API ───────────────────────────────────────────────────────────────

window.intercomMonitor = {
  init:           initIntercomMonitor,
  updateStations: updateIntercomStations,
  render:         renderMonitorSection,
  handle:         handleIntercomMonitorAction,
  destroy:        destroyIntercomMonitor,
  get listeningTo()  { return _m.listeningTo; },
  get reqStatus()    { return _m.reqStatus; }
};

})();
