/**
 * intercom-field.js
 * AgriX PWA — Otkupac intercom field module (Phase 1, auto-activate)
 *
 * Design contract:
 *   - Zero interaction after mic permission is granted once at setup
 *   - Bootstraps a Firebase onValue listener on /intercom/requests/{entityID}
 *   - On status="requested": getUserMedia (no gesture needed, permission pre-granted)
 *     → MediaRecorder 500ms chunks → Firebase /sessions/{entityID}/toMgmt
 *   - On status="ended": mic stops, silent cleanup
 *   - Visual: single small 🔴 indicator injected into app header when live
 *
 * Permission setup (one-time, from user gesture):
 *   window.intercomField.requestMicPermission()  — tab-queue device setup
 *   window.intercomField.micPermissionState()    — 'granted'|'prompt'|'denied'
 *   window.intercomField.renderPermissionStatus() — refresh status UI row
 *
 * Dependencies (window globals):
 *   CONFIG     — ENTITY_ID, ENTITY_NAME, STATION_ID,
 *                FIREBASE_API_KEY, FIREBASE_PROJECT_ID,
 *                FIREBASE_APP_ID, FIREBASE_RTDB_URL
 *   showToast, byId, showEl, hideEl, setText  — existing utils
 *   firebase   — Firebase compat UMD (app + auth + database)
 *
 * NOT in scope (Phase 1):
 *   - fromMgmt back-channel (Management → Otkupac audio)
 *   - Custom token auth (Anonymous Auth used)
 *   - Push notifications for background wake
 */

'use strict';

(function () {

// ─── Module state ─────────────────────────────────────────────────────────────

const _f = {
  db:            null,
  requestUnsub:  null,   // Firebase listener detach fn
  stream:        null,
  mediaRecorder: null,
  sessionRef:    null,
  seq:           0,
  status:        'idle'  // 'idle' | 'activating' | 'active' | 'stopping'
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

  _f.db = app.database();

  if (!_f.db) {
    throw new Error('app.database() vratio null/undefined');
  }
}

async function _ensureAuth() {
  const auth = firebase.app('agrix-intercom').auth();
  if (auth.currentUser) return;
  await auth.signInAnonymously();
}

// ─── Bootstrap — called from app.js Otkupac role bootstrap ───────────────────

async function initIntercomField() {
  try {
    _initFirebase();
    await _ensureAuth();
    _startRequestListener();
    _renderHeaderIndicator();
    _renderPermissionStatus();   // refresh setup row if visible
  } catch (err) {
    console.error('[intercom-field] init failed:', err);
    // Non-fatal — otkup flow continues without intercom
  }
}

// ─── Request listener ─────────────────────────────────────────────────────────

function _startRequestListener() {
  if (_f.requestUnsub) return;

  const reqRef = _f.db.ref(`intercom/requests/${CONFIG.ENTITY_ID}`);

  const handler = reqRef.on('value', snapshot => {
    const data = snapshot.val();
    if (!data) return;

    switch (data.status) {
      case 'requested':
        if (_f.status === 'idle') _activateMic();
        break;
      case 'ended':
        if (_f.status === 'active' || _f.status === 'activating') _deactivateMic();
        break;
    }
  });

  _f.requestUnsub = () => reqRef.off('value', handler);
}

// ─── Mic activation (automatic — no gesture required after permission granted) -

async function _activateMic() {
  _f.status = 'activating';

  try {
    // Permission must already be 'granted' from one-time setup
    _f.stream = await navigator.mediaDevices.getUserMedia({
      audio: {
        echoCancellation: true,
        noiseSuppression: true,
        sampleRate: 16000   // lower rate = smaller chunks, adequate for voice
      },
      video: false
    });

    // Write session meta
    _f.sessionRef = _f.db.ref(`intercom/sessions/${CONFIG.ENTITY_ID}`);
    await _f.sessionRef.child('meta').set({
      name:      CONFIG.ENTITY_NAME || CONFIG.ENTITY_ID,
      stationID: CONFIG.STATION_ID  || '',
      startedAt: firebase.database.ServerValue.TIMESTAMP,
      status:    'active'
    });

    // onDisconnect: clean up if device loses connection mid-session
    _f.sessionRef.child('meta/status').onDisconnect().set('ended');
    _f.db.ref(`intercom/requests/${CONFIG.ENTITY_ID}/status`)
      .onDisconnect().set('ended');

    // Confirm to Management that mic is live
    await _f.db.ref(`intercom/requests/${CONFIG.ENTITY_ID}/status`).set('active');

    // Start recording
    const mimeType = _bestMimeType();
    _f.mediaRecorder = new MediaRecorder(_f.stream, { mimeType });
    _f.seq = 0;

    _f.mediaRecorder.ondataavailable = _onChunk;
    _f.mediaRecorder.onerror = () => _deactivateMic();
    _f.mediaRecorder.start(500);

    _f.status = 'active';
    _setHeaderIndicator(true);

  } catch (err) {
    _f.status = 'idle';
    _setHeaderIndicator(false);

    if (err.name === 'NotAllowedError') {
      // Permission revoked in browser settings after setup
      showToast('Mikrofon nije dostupan — proveri dozvole', 'error');
    } else {
      console.error('[intercom-field] _activateMic error:', err);
    }

    // Tell Management activation failed so UI doesn't hang on "waiting"
    _f.db.ref(`intercom/requests/${CONFIG.ENTITY_ID}/status`)
      .set('error').catch(() => {});
  }
}

// ─── Mic deactivation ─────────────────────────────────────────────────────────

async function _deactivateMic() {
  if (_f.status === 'idle' || _f.status === 'stopping') return;
  _f.status = 'stopping';

  try {
    if (_f.mediaRecorder?.state !== 'inactive') {
      _f.mediaRecorder.stop();
    }
    _f.stream?.getTracks().forEach(t => t.stop());

    if (_f.sessionRef) {
      await _f.sessionRef.child('meta/status').set('ended');
      _f.sessionRef.child('meta/status').onDisconnect().cancel();
      _f.db.ref(`intercom/requests/${CONFIG.ENTITY_ID}/status`)
        .onDisconnect().cancel();
    }
  } catch (err) {
    console.warn('[intercom-field] _deactivateMic error:', err);
  } finally {
    _f.stream        = null;
    _f.mediaRecorder = null;
    _f.sessionRef    = null;
    _f.seq           = 0;
    _f.status        = 'idle';
    _setHeaderIndicator(false);
  }
}

// ─── Chunk write ──────────────────────────────────────────────────────────────

async function _onChunk(e) {
  if (!e.data || e.data.size === 0 || _f.status !== 'active') return;
  try {
    const b64 = await _toBase64(e.data);
    await _f.sessionRef.child('toMgmt').set({
      seq:   ++_f.seq,
      chunk: b64,
      ts:    Date.now()
    });
  } catch (err) {
    console.warn('[intercom-field] chunk write failed:', err);
  }
}

// ─── Header indicator ─────────────────────────────────────────────────────────
// Single small element appended to existing #headerInfo — no layout changes.

function _renderHeaderIndicator() {
  if (byId('intercom-header-dot')) return;
  const anchor = byId('headerInfo') || document.querySelector('.app-header');
  if (!anchor) return;
  const dot = document.createElement('span');
  dot.id        = 'intercom-header-dot';
  dot.className = 'intercom-header-dot';
  dot.title     = 'Interkom aktivan — Dispečer sluša';
  dot.textContent = '🔴';
  dot.style.display = 'none';
  anchor.appendChild(dot);
}

function _setHeaderIndicator(active) {
  const dot = byId('intercom-header-dot');
  if (!dot) return;
  dot.style.display = active ? 'inline' : 'none';
}

// ─── One-time mic permission setup ───────────────────────────────────────────
// Must be called from a user gesture (button click).

async function requestMicPermission() {
  try {
    const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
    stream.getTracks().forEach(t => t.stop());
    showToast('Mikrofon dozvoljen ✓', 'success');
    _renderPermissionStatus();
    return true;
  } catch {
    showToast('Mikrofon odbijen — proveri podešavanja browsera', 'error');
    return false;
  }
}

async function micPermissionState() {
  try {
    const result = await navigator.permissions.query({ name: 'microphone' });
    return result.state;   // 'granted' | 'prompt' | 'denied'
  } catch {
    return 'prompt';   // Permissions API not supported
  }
}

async function _renderPermissionStatus() {
  const statusEl = byId('intercom-permission-status');
  const btnEl    = byId('btn-intercom-permission');
  if (!statusEl) return;

  const state = await micPermissionState();

  if (state === 'granted') {
    statusEl.innerHTML = `<span class="intercom-perm-ok">✓ Mikrofon dozvoljen</span>`;
    if (btnEl) btnEl.style.display = 'none';
  } else if (state === 'denied') {
    statusEl.innerHTML = `<span class="intercom-perm-denied">✗ Mikrofon blokiran — omogući u podešavanjima browsera</span>`;
    if (btnEl) btnEl.style.display = 'none';
  } else {
    statusEl.innerHTML = `<span class="intercom-perm-prompt">Mikrofon nije podešen</span>`;
    if (btnEl) btnEl.style.display = 'inline-block';
  }
}

// ─── data-action handler (routed from app.js handleAppShellClick) ─────────────

function handleIntercomFieldAction(action) {
  if (action === 'intercom-request-permission') {
    requestMicPermission();
  }
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function _toBase64(blob) {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onloadend = () => resolve(r.result.split(',')[1]);
    r.onerror   = reject;
    r.readAsDataURL(blob);
  });
}

function _bestMimeType() {
  const candidates = [
    'audio/webm;codecs=opus',
    'audio/webm',
    'audio/ogg;codecs=opus',
    ''
  ];
  return candidates.find(t => !t || MediaRecorder.isTypeSupported(t)) || '';
}

// ─── Page unload ──────────────────────────────────────────────────────────────

window.addEventListener('beforeunload', () => {
  if (_f.status === 'active') {
    _f.mediaRecorder?.stop();
    _f.stream?.getTracks().forEach(t => t.stop());
    // Firebase onDisconnect handles meta/status and request/status cleanup
  }
});

// ─── Public API ───────────────────────────────────────────────────────────────

window.intercomField = {
  init:                   initIntercomField,
  handle:                 handleIntercomFieldAction,
  requestMicPermission:   requestMicPermission,
  micPermissionState:     micPermissionState,
  renderPermissionStatus: _renderPermissionStatus,
  get status() { return _f.status; }
};

})();
