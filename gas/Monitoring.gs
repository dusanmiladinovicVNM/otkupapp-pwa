/*
============================================================
OtkupApp / AgriX - Monitoring.gs
============================================================

Napomena:
- PWA monitoring treba da ide kroz postojeći token.
- VBA monitoring može da ide kroz MONITORING_INGEST_SECRET ako VBA nema PWA token.
- Ne loguju se tokeni, passwordi, SEF API ključevi, puni XML/PDF payload-i.
*/

// ============================================================
// CONFIG
// ============================================================

const MONITORING_SPREADSHEET_NAME = 'OtkupApp_Monitoring_PROD';
// Kanonski Script Property kljuc (isti naming kao alert/ingest ispod, kao u docs/ARCHITECTURE_REFERENCE.md).
const MONITORING_PROP_SPREADSHEET_ID = 'MONITORING_SPREADSHEET_ID';
// Legacy kljuc koji je stara verzija koda citala/sugerisala u error poruci; citamo ga kao fallback radi kompatibilnosti.
const MONITORING_PROP_SPREADSHEET_ID_LEGACY = 'MONITORING_PROP_SPREADSHEET_ID';
const MONITORING_PROP_ALERT_EMAIL = 'MONITORING_ALERT_EMAIL';
const MONITORING_PROP_INGEST_SECRET = 'MONITORING_INGEST_SECRET';

const MONITORING_MAX_PAYLOAD_CHARS = 2000;
const MONITORING_MAX_MESSAGE_CHARS = 600;
const MONITORING_MAX_STACK_CHARS = 2000;

const MONITORING_TABS = {
  Health: [
    'Component',
    'Status',
    'LastOK',
    'LastError',
    'ErrorCount24h',
    'WarningCount24h',
    'CriticalCount24h',
    'LastMessage',
    'OwnerAction',
    'UpdatedAt'
  ],
  Events: [
    'Timestamp',
    'Environment',
    'Source',
    'UserId',
    'Role',
    'DeviceId',
    'AppVersion',
    'EventType',
    'EntityType',
    'EntityId',
    'CorrelationId',
    'Severity',
    'Message',
    'PayloadJson',
    'BuildSha',
    'BuildDate',
    'BuildVersion'
  ],
  Errors: [
    'Timestamp',
    'Environment',
    'Source',
    'Severity',
    'UserId',
    'DeviceId',
    'AppVersion',
    'Module',
    'FunctionName',
    'ErrorCode',
    'ErrorMessage',
    'StackTrace',
    'EntityType',
    'EntityId',
    'CorrelationId',
    'Resolved',
    'ResolvedAt',
    'ResolutionNote'
  ],
  SyncStatus: [
    'Timestamp',
    'UserId',
    'DeviceId',
    'AppVersion',
    'OnlineStatus',
    'LastSyncAt',
    'PendingCount',
    'FailedCount',
    'ConflictCount',
    'OldestPendingAgeMin',
    'LastSyncDurationMs',
    'LastSyncResult',
    'LastError'
  ],
  SEFStatus: [
    'Timestamp',
    'InvoiceLocalId',
    'BusinessInvoiceNo',
    'SEFRequestId',
    'SEFInvoiceId',
    'SEFStatus',
    'LocalStatus',
    'LastAttemptAt',
    'AttemptCount',
    'LastHttpCode',
    'LastError',
    'NextAction',
    'CorrelationId',
    'NeedsManualReview'
  ],
  UserSessions: [
    'Timestamp',
    'UserId',
    'Role',
    'DeviceId',
    'AppVersion',
    'BuildNumber',
    'Browser',
    'OS',
    'IsOnline',
    'LastSeenAt',
    'LastRoute'
  ],
  Backups: [
    'Timestamp',
    'BackupType',
    'SourceSpreadsheetId',
    'BackupFileId',
    'BackupLocation',
    'RowsCount',
    'Checksum',
    'Status',
    'DurationMs',
    'ErrorMessage'
  ],
  Alerts: [
    'Timestamp',
    'Severity',
    'Component',
    'AlertCode',
    'Message',
    'CorrelationId',
    'SentTo',
    'SentAt',
    'Acknowledged',
    'AcknowledgedBy',
    'Resolved',
    'ResolvedAt'
  ],
  AuditCritical: [
    'Timestamp',
    'Source',
    'UserId',
    'Role',
    'Action',
    'EntityType',
    'EntityId',
    'CorrelationId',
    'BeforeHash',
    'AfterHash',
    'Message',
    'PayloadJson'
  ],
  Fleet: [
    'DeviceId',
    'UserId',
    'Source',
    'AppVersion',
    'BuildVersion',
    'BuildSha',
    'BuildDate',
    'LastSeenAt',
    'FirstSeenAt',
    'EventCount'
  ]
};

const MONITORING_COMPONENTS = [
  'PWA Sync',
  'GAS API',
  'Google Sheets DB',
  'VBA Client',
  'SEF API',
  'Backup',
  'Auth',
  'MasterData Sync',
  'Offline Queue'
];

// ============================================================
// ROUTER HANDLERS
// ============================================================

function handlePublicMonitoringAction(data) {
  const action = String(data && data.action || '');
  if (action !== 'monitorPublic') return null;

  if (!isValidMonitoringSecret_(data.monitoringSecret)) {
    return { success: false, error: 'Invalid monitoring secret', code: 401 };
  }

  const result = monitoringIngest_(data, {
    userId: data.userId || '',
    role: data.role || 'System',
    entityID: data.entityId || data.entityID || ''
  });

  return result;
}

// Min-version gate (public). VBA klijent salje svoju verziju; vracamo policy
// iz Script Properties (menja se BEZ redeploy-a). Autentikacija = isti
// MONITORING_INGEST_SECRET kao monitorPublic.
//   VERSION_MIN      minimalna dozvoljena (npr. "2.2.0"); prazno = gate iskljucen
//   VERSION_LATEST   najnovija (samo za poruku korisniku)
//   VERSION_ENFORCE  YES/NO (default NO = klijent samo upozori, ne blokira)
//   VERSION_MESSAGE  opciono custom tekst
function handleCheckVersionPublic_(data) {
  if (!isValidMonitoringSecret_(data && data.monitoringSecret)) {
    return { success: false, error: 'Invalid monitoring secret', code: 401 };
  }

  const props = PropertiesService.getScriptProperties();
  const enforceRaw = String(props.getProperty('VERSION_ENFORCE') || '').trim().toUpperCase();
  const enforce = (enforceRaw === 'YES' || enforceRaw === 'TRUE' || enforceRaw === '1' || enforceRaw === 'ON');

  return {
    success: true,
    minVersion: String(props.getProperty('VERSION_MIN') || '').trim(),
    latestVersion: String(props.getProperty('VERSION_LATEST') || '').trim(),
    enforce: enforce ? 'YES' : 'NO',
    message: String(props.getProperty('VERSION_MESSAGE') || '').trim()
  };
}

function handleMonitoringAction(data, tokenData) {
  const action = String(data && data.action || '');

  if (action === 'monitor') {
    return monitoringIngest_(data, tokenData || {});
  }

  if (action === 'getMonitoringHealth') {
    if (!monitoringCanReadHealth_(tokenData)) {
      return { success: false, error: 'Nemate pristup', code: 403 };
    }
    return getMonitoringHealth();
  }

  if (action === 'getMonitoringFleet') {
    if (!monitoringCanReadHealth_(tokenData)) {
      return { success: false, error: 'Nemate pristup', code: 403 };
    }
    return getMonitoringFleet();
  }

  if (action === 'rebuildMonitoringFleet') {
    if (!monitoringCanReadHealth_(tokenData)) {
      return { success: false, error: 'Nemate pristup', code: 403 };
    }
    return rebuildMonitoringFleet();
  }

  if (action === 'ackMonitoringAlert') {
    if (!monitoringCanReadHealth_(tokenData)) {
      return { success: false, error: 'Nemate pristup', code: 403 };
    }
    return acknowledgeMonitoringAlert(data, tokenData);
  }

  if (action === 'resolveMonitoringAlert') {
    if (!monitoringCanReadHealth_(tokenData)) {
      return { success: false, error: 'Nemate pristup', code: 403 };
    }
    return resolveMonitoringAlert(data, tokenData);
  }

  return null;
}

function monitoringCanReadHealth_(tokenData) {
  if (!tokenData) return false;
  const role = String(tokenData.role || '').toLowerCase();
  return role === 'management' || role === 'admin' || role === 'operator';
}

// ============================================================
// MAIN INGEST
// ============================================================

function monitoringIngest_(input, tokenData) {
  const now = new Date();
  const event = normalizeMonitoringEvent_(input, tokenData, now);

  return monitoringWithLock_(function() {
    const ss = getMonitoringSpreadsheet_();
    ensureMonitoringWorkbook_(ss);

    appendMonitoringEvent_(ss, event);

    if (event.severity === 'ERROR' || event.severity === 'CRITICAL') {
      appendMonitoringError_(ss, event);
    }

    if (isSyncEvent_(event)) {
      appendSyncStatus_(ss, event);
    }

    if (isSessionEvent_(event)) {
      appendUserSession_(ss, event);
    }

    if (isSEFEvent_(event)) {
      upsertSEFStatus_(ss, event);
    }

    if (isBackupEvent_(event)) {
      appendBackupStatus_(ss, event);
    }

    if (isCriticalAuditEvent_(event)) {
      appendAuditCritical_(ss, event);
    }

    updateComponentHealthFromEvent_(ss, event);
    maybeCreateAlertFromEvent_(ss, event);

    return {
      success: true,
      eventId: event.eventId,
      timestamp: event.timestamp,
      severity: event.severity,
      component: event.component
    };
  });
}

function normalizeMonitoringEvent_(input, tokenData, now) {
  const payload = sanitizePayload_(input.payload || input.payloadJson || {});
  const source = normalizeSource_(input.source || payload.source || 'GAS');
  const eventType = String(input.eventType || input.errorAction || payload.eventType || 'EVENT');
  const severity = normalizeSeverity_(input.severity || payload.severity || 'INFO');

  const tokenUserId = tokenData && (tokenData.userId || tokenData.username || tokenData.entityID || tokenData.entityId) || '';
  const tokenRole = tokenData && tokenData.role || '';

  const correlationId = String(
    input.correlationId ||
    input.CorrelationId ||
    payload.correlationId ||
    input.entityId ||
    input.entityID ||
    payload.entityId ||
    ''
  );

  return {
    eventId: Utilities.getUuid(),
    timestamp: now.toISOString(),
    environment: String(input.environment || input.env || 'PROD'),
    source: source,
    userId: String(input.userId || payload.userId || tokenUserId || ''),
    role: String(input.role || payload.role || tokenRole || ''),
    deviceId: String(input.deviceId || input.DeviceId || payload.deviceId || ''),
    appVersion: String(input.appVersion || payload.appVersion || ''),
    buildSha: String(input.buildSha || payload.buildSha || ''),
    buildVersion: String(input.buildVersion || payload.buildVersion || ''),
    buildDate: String(input.buildDate || payload.buildDate || ''),
    buildNumber: String(input.buildNumber || payload.buildNumber || ''),
    browser: String(input.browser || payload.browser || ''),
    os: String(input.os || payload.os || ''),
    eventType: eventType,
    entityType: String(input.entityType || payload.entityType || ''),
    entityId: String(input.entityId || input.entityID || payload.entityId || ''),
    correlationId: correlationId,
    severity: severity,
    component: inferComponent_(source, eventType, input.component || payload.component),
    module: String(input.module || payload.module || ''),
    functionName: String(input.functionName || input.procedure || payload.functionName || ''),
    errorCode: String(input.errorCode || payload.errorCode || ''),
    message: truncate_(String(input.message || input.errorMessage || payload.message || ''), MONITORING_MAX_MESSAGE_CHARS),
    stackTrace: truncate_(String(input.stackTrace || input.stack || payload.stackTrace || ''), MONITORING_MAX_STACK_CHARS),
    payload: payload,
    raw: input
  };
}

// ============================================================
// WRITERS
// ============================================================

function appendMonitoringEvent_(ss, event) {
  const sh = getMonitoringSheet_(ss, 'Events');
  sh.appendRow([
    event.timestamp,
    event.environment,
    event.source,
    event.userId,
    event.role,
    event.deviceId,
    event.appVersion,
    event.eventType,
    event.entityType,
    event.entityId,
    event.correlationId,
    event.severity,
    event.message,
    stringifyPayload_(event.payload),
    event.buildSha,
    event.buildDate,
    event.buildVersion
  ]);
}

function appendMonitoringError_(ss, event) {
  const sh = getMonitoringSheet_(ss, 'Errors');
  sh.appendRow([
    event.timestamp,
    event.environment,
    event.source,
    event.severity,
    event.userId,
    event.deviceId,
    event.appVersion,
    event.module,
    event.functionName,
    event.errorCode,
    event.message,
    event.stackTrace,
    event.entityType,
    event.entityId,
    event.correlationId,
    false,
    '',
    ''
  ]);
}

function appendSyncStatus_(ss, event) {
  const p = event.payload || {};
  const sh = getMonitoringSheet_(ss, 'SyncStatus');
  sh.appendRow([
    event.timestamp,
    event.userId,
    event.deviceId,
    event.appVersion,
    boolToText_(p.online !== undefined ? p.online : p.onlineStatus),
    String(p.lastSyncAt || ''),
    numberOrZero_(p.pendingCount),
    numberOrZero_(p.failedCount),
    numberOrZero_(p.conflictCount),
    numberOrZero_(p.oldestPendingAgeMin),
    numberOrZero_(p.lastSyncDurationMs),
    String(p.lastSyncResult || event.eventType),
    String(p.lastError || event.message || '')
  ]);
}

function appendUserSession_(ss, event) {
  const p = event.payload || {};
  const sh = getMonitoringSheet_(ss, 'UserSessions');
  sh.appendRow([
    event.timestamp,
    event.userId,
    event.role,
    event.deviceId,
    event.appVersion,
    event.buildNumber,
    event.browser,
    event.os,
    boolToText_(p.online !== undefined ? p.online : p.isOnline),
    String(p.lastSeenAt || event.timestamp),
    String(p.route || p.lastRoute || '')
  ]);
}

function upsertSEFStatus_(ss, event) {
  const p = event.payload || {};
  const sh = getMonitoringSheet_(ss, 'SEFStatus');
  const correlationId = event.correlationId || p.correlationId || p.invoiceLocalId || p.businessInvoiceNo || '';
  const invoiceLocalId = String(p.invoiceLocalId || event.entityId || '');

  const rowIndex = findSEFStatusRow_(sh, correlationId, invoiceLocalId);
  const row = [
    event.timestamp,
    invoiceLocalId,
    String(p.businessInvoiceNo || ''),
    String(p.sefRequestId || p.SEFRequestId || ''),
    String(p.sefInvoiceId || p.SEFInvoiceId || ''),
    String(p.sefStatus || p.SEFStatus || inferSEFStatusFromEvent_(event)),
    String(p.localStatus || p.LocalStatus || ''),
    String(p.lastAttemptAt || event.timestamp),
    numberOrZero_(p.attemptCount),
    String(p.lastHttpCode || p.httpCode || ''),
    String(p.lastError || event.message || ''),
    String(p.nextAction || inferSEFNextAction_(event)),
    correlationId,
    boolToText_(p.needsManualReview !== undefined ? p.needsManualReview : event.severity === 'CRITICAL')
  ];

  if (rowIndex > 0) {
    sh.getRange(rowIndex, 1, 1, row.length).setValues([row]);
  } else {
    sh.appendRow(row);
  }
}

function appendBackupStatus_(ss, event) {
  const p = event.payload || {};
  const sh = getMonitoringSheet_(ss, 'Backups');
  sh.appendRow([
    event.timestamp,
    String(p.backupType || event.entityType || ''),
    String(p.sourceSpreadsheetId || ''),
    String(p.backupFileId || ''),
    String(p.backupLocation || ''),
    numberOrZero_(p.rowsCount),
    String(p.checksum || ''),
    String(p.status || inferBackupStatusFromEvent_(event)),
    numberOrZero_(p.durationMs),
    String(p.errorMessage || event.message || '')
  ]);
}

function appendAuditCritical_(ss, event) {
  const p = event.payload || {};
  const sh = getMonitoringSheet_(ss, 'AuditCritical');
  sh.appendRow([
    event.timestamp,
    event.source,
    event.userId,
    event.role,
    event.eventType,
    event.entityType,
    event.entityId,
    event.correlationId,
    String(p.beforeHash || ''),
    String(p.afterHash || ''),
    event.message,
    stringifyPayload_(event.payload)
  ]);
}

// ============================================================
// HEALTH
// ============================================================

function updateComponentHealthFromEvent_(ss, event) {
  const component = event.component || 'GAS API';
  const status = statusFromSeverity_(event.severity);
  upsertHealthRow_(ss, component, status, event);

  if (event.source === 'PWA' && isSyncEvent_(event)) {
    updateOfflineQueueHealth_(ss, event);
  }
}

function upsertHealthRow_(ss, component, status, event) {
  const sh = getMonitoringSheet_(ss, 'Health');
  const idx = findRowByFirstColumn_(sh, component);
  const counts = countRecentEventsByComponent_(ss, component, 24);

  let currentStatus = status;
  if (idx > 0) {
    const existing = String(sh.getRange(idx, 2).getValue() || 'UNKNOWN');
    currentStatus = worseStatus_(existing, status);
  }

  const isOk = currentStatus === 'OK';
  const row = [
    component,
    currentStatus,
    isOk ? event.timestamp : getExistingCell_(sh, idx, 3),
    !isOk ? event.timestamp : getExistingCell_(sh, idx, 4),
    counts.ERROR,
    counts.WARN,
    counts.CRITICAL,
    event.message || event.eventType,
    ownerActionForStatus_(component, currentStatus, event),
    event.timestamp
  ];

  if (idx > 0) {
    sh.getRange(idx, 1, 1, row.length).setValues([row]);
  } else {
    sh.appendRow(row);
  }
}

function updateOfflineQueueHealth_(ss, event) {
  const p = event.payload || {};
  const pending = numberOrZero_(p.pendingCount);
  const failed = numberOrZero_(p.failedCount);
  const conflicts = numberOrZero_(p.conflictCount);
  const oldest = numberOrZero_(p.oldestPendingAgeMin);

  let status = 'OK';
  let message = 'Offline queue OK';

  if (conflicts > 0 || oldest > 30 || failed > 0) {
    status = 'CRITICAL';
    message = 'Offline queue needs manual review: pending=' + pending + ', failed=' + failed + ', conflicts=' + conflicts + ', oldestMin=' + oldest;
  } else if (pending > 0) {
    status = 'WARN';
    message = 'Offline queue has pending records: ' + pending;
  }

  upsertHealthRow_(ss, 'Offline Queue', status, {
    timestamp: event.timestamp,
    severity: status === 'OK' ? 'INFO' : status,
    eventType: event.eventType,
    message: message,
    component: 'Offline Queue'
  });
}

function getMonitoringHealth() {
  return monitoringWithLock_(function() {
    const ss = getMonitoringSpreadsheet_();
    ensureMonitoringWorkbook_(ss);
    const health = readSheetAsObjects_(getMonitoringSheet_(ss, 'Health'));
    const alerts = readOpenAlerts_(ss);
    return {
      success: true,
      timestamp: new Date().toISOString(),
      health: health,
      openAlerts: alerts
    };
  });
}

// ============================================================
// FLEET INVENTORY ("ko ima koju verziju")
// Agregat poslednjeg stanja po uredjaju iz Events taba.
// rebuildMonitoringFleet() moze se pokrenuti rucno iz Apps Script editora.
// ============================================================

function rebuildMonitoringFleet() {
  return monitoringWithLock_(function() {
    const ss = getMonitoringSpreadsheet_();
    ensureMonitoringWorkbook_(ss);
    const rows = buildFleetRows_(ss);
    writeFleetSheet_(ss, rows);
    return { success: true, devices: rows.length, rebuiltAt: new Date().toISOString() };
  });
}

function getMonitoringFleet() {
  return monitoringWithLock_(function() {
    const ss = getMonitoringSpreadsheet_();
    ensureMonitoringWorkbook_(ss);
    return { success: true, timestamp: new Date().toISOString(), fleet: buildFleetRows_(ss) };
  });
}

function buildFleetRows_(ss) {
  const events = readSheetAsObjects_(getMonitoringSheet_(ss, 'Events'));
  const byDevice = {};

  events.forEach(function(e) {
    const deviceId = String(e.DeviceId || '').trim();
    if (!deviceId) return;
    const ts = fleetTs_(e.Timestamp);

    let d = byDevice[deviceId];
    if (!d) {
      d = byDevice[deviceId] = {
        deviceId: deviceId,
        userId: String(e.UserId || ''),
        source: String(e.Source || ''),
        appVersion: String(e.AppVersion || ''),
        buildVersion: String(e.BuildVersion || ''),
        buildSha: String(e.BuildSha || ''),
        buildDate: String(e.BuildDate || ''),
        lastSeenAt: ts,
        firstSeenAt: ts,
        count: 0
      };
    }

    d.count++;
    if (ts && ts >= d.lastSeenAt) {
      d.lastSeenAt = ts;
      d.userId = String(e.UserId || d.userId);
      d.source = String(e.Source || d.source);
      d.appVersion = String(e.AppVersion || d.appVersion);
      d.buildVersion = String(e.BuildVersion || d.buildVersion);
      d.buildSha = String(e.BuildSha || d.buildSha);
      d.buildDate = String(e.BuildDate || d.buildDate);
    }
    if (ts && (!d.firstSeenAt || ts < d.firstSeenAt)) d.firstSeenAt = ts;
  });

  return Object.keys(byDevice)
    .map(function(k) { return byDevice[k]; })
    .sort(function(a, b) { return a.lastSeenAt < b.lastSeenAt ? 1 : -1; });
}

function writeFleetSheet_(ss, rows) {
  const sh = getMonitoringSheet_(ss, 'Fleet');
  ensureHeader_(sh, MONITORING_TABS.Fleet);

  const last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, sh.getLastColumn()).clearContent();
  if (!rows.length) return;

  const values = rows.map(function(d) {
    return [d.deviceId, d.userId, d.source, d.appVersion, d.buildVersion, d.buildSha, d.buildDate, d.lastSeenAt, d.firstSeenAt, d.count];
  });
  sh.getRange(2, 1, values.length, MONITORING_TABS.Fleet.length).setValues(values);
}

function fleetTs_(v) {
  if (v instanceof Date) return v.toISOString();
  return String(v || '');
}

// ============================================================
// ALERTS
// ============================================================

function maybeCreateAlertFromEvent_(ss, event) {
  const p = event.payload || {};
  const component = event.component || 'GAS API';
  let alertCode = '';
  let severity = event.severity;
  let message = event.message || event.eventType;

  if (event.severity === 'CRITICAL') {
    alertCode = componentKey_(component) + '_CRITICAL';
  }

  if (isSEFEvent_(event)) {
    const sefStatus = String(p.sefStatus || p.SEFStatus || '').toUpperCase();
    const nextAction = String(p.nextAction || '').toUpperCase();
    if (sefStatus === 'UNKNOWN' || nextAction === 'MANUAL_REVIEW') {
      alertCode = 'SEF_UNKNOWN_MANUAL_REVIEW';
      severity = 'CRITICAL';
      message = message || 'SEF status UNKNOWN / manual review needed';
    }
    if (sefStatus === 'STUCK') {
      alertCode = 'SEF_STUCK';
      severity = 'CRITICAL';
      message = message || 'SEF invoice stuck';
    }
  }

  if (isSyncEvent_(event)) {
    if (numberOrZero_(p.failedCount) > 0) {
      alertCode = 'SYNC_FAILED_DEVICE';
      severity = worseSeverity_(severity, 'WARN');
      message = 'Sync failed on device ' + event.deviceId + ': failed=' + numberOrZero_(p.failedCount);
    }
    if (numberOrZero_(p.oldestPendingAgeMin) > 30 || numberOrZero_(p.conflictCount) > 0) {
      alertCode = 'SYNC_PENDING_OR_CONFLICT';
      severity = 'CRITICAL';
      message = 'Old pending/conflict on device ' + event.deviceId + ': oldestMin=' + numberOrZero_(p.oldestPendingAgeMin) + ', conflicts=' + numberOrZero_(p.conflictCount);
    }
  }

  if (isBackupEvent_(event)) {
    const backupStatus = String(p.status || '').toUpperCase();
    if (backupStatus === 'FAILED' || event.eventType === 'BACKUP_FAIL') {
      alertCode = 'BACKUP_FAILED';
      severity = 'CRITICAL';
      message = message || 'Backup failed';
    }
  }

  if (!alertCode) return;
  createMonitoringAlert_(ss, severity, component, alertCode, message, event.correlationId || event.deviceId || event.entityId || '');
}

function createMonitoringAlert_(ss, severity, component, alertCode, message, correlationId) {
  const sh = getMonitoringSheet_(ss, 'Alerts');
  if (openAlertExists_(sh, component, alertCode, correlationId)) return;

  const now = new Date().toISOString();
  const alertEmail = getMonitoringAlertEmail_();
  let sentAt = '';

  if (alertEmail && (severity === 'CRITICAL' || severity === 'ERROR')) {
    try {
      MailApp.sendEmail({
        to: alertEmail,
        subject: '[OtkupApp Monitoring] ' + severity + ' - ' + alertCode,
        body:
          'Component: ' + component + '\n' +
          'Alert: ' + alertCode + '\n' +
          'Severity: ' + severity + '\n' +
          'CorrelationId: ' + correlationId + '\n' +
          'Message: ' + message + '\n' +
          'Time: ' + now
      });
      sentAt = now;
    } catch (err) {
      Logger.log('Monitoring alert email failed: ' + err.message);
    }
  }

  sh.appendRow([
    now,
    severity,
    component,
    alertCode,
    truncate_(message, MONITORING_MAX_MESSAGE_CHARS),
    correlationId,
    alertEmail,
    sentAt,
    false,
    '',
    false,
    ''
  ]);
}

function acknowledgeMonitoringAlert(data, tokenData) {
  return updateMonitoringAlertState_(data, tokenData, 'ack');
}

function resolveMonitoringAlert(data, tokenData) {
  return updateMonitoringAlertState_(data, tokenData, 'resolve');
}

function updateMonitoringAlertState_(data, tokenData, mode) {
  return monitoringWithLock_(function() {
    const ss = getMonitoringSpreadsheet_();
    const sh = getMonitoringSheet_(ss, 'Alerts');
    const rowIndex = numberOrZero_(data.rowIndex);
    if (rowIndex < 2) return { success: false, error: 'Invalid rowIndex' };

    const user = String(tokenData.username || tokenData.userId || tokenData.entityID || 'operator');
    const now = new Date().toISOString();

    if (mode === 'ack') {
      sh.getRange(rowIndex, 9).setValue(true);
      sh.getRange(rowIndex, 10).setValue(user);
    }

    if (mode === 'resolve') {
      sh.getRange(rowIndex, 11).setValue(true);
      sh.getRange(rowIndex, 12).setValue(now);
    }

    return { success: true, rowIndex: rowIndex, mode: mode };
  });
}

// ============================================================
// WATCHDOG / SCHEDULED JOBS
// ============================================================

function runMonitoringWatchdog() {
  return monitoringWithLock_(function() {
    const ss = getMonitoringSpreadsheet_();
    ensureMonitoringWorkbook_(ss);
    const now = new Date();

    runMonitoringActiveHealthChecks_(ss, now);

    checkSEFStuck_(ss, now);
    checkSyncStuck_(ss, now);
    checkBackupFreshness_(ss, now);
    checkRecentGasErrors_(ss, now);

    return { success: true, checkedAt: now.toISOString() };
  });
}

function sendDailyMonitoringSummary() {
  const email = getMonitoringAlertEmail_();
  if (!email) return { success: false, error: 'MONITORING_ALERT_EMAIL not set' };

  const result = getMonitoringHealth();
  if (!result.success) return result;

  const lines = [];
  lines.push('Production Health Summary');
  lines.push('Time: ' + result.timestamp);
  lines.push('');
  lines.push('Health:');

  (result.health || []).forEach(function(row) {
    lines.push('- ' + row.Component + ': ' + row.Status + ' | ' + (row.LastMessage || ''));
  });

  lines.push('');
  lines.push('Open alerts: ' + (result.openAlerts || []).length);
  (result.openAlerts || []).slice(0, 20).forEach(function(a) {
    lines.push('- ' + a.Severity + ' ' + a.Component + ' ' + a.AlertCode + ': ' + a.Message);
  });

  MailApp.sendEmail({
    to: email,
    subject: '[OtkupApp Monitoring] Daily Health Summary',
    body: lines.join('\n')
  });

  return { success: true, sentTo: email };
}

function installMonitoringTriggers() {
  removeMonitoringTriggers_();

  ScriptApp.newTrigger('runMonitoringWatchdog')
    .timeBased()
    .everyMinutes(15)
    .create();

  ScriptApp.newTrigger('sendDailyMonitoringSummary')
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .create();

  ScriptApp.newTrigger('rebuildMonitoringFleet')
    .timeBased()
    .everyHours(1)
    .create();

  return { success: true };
}

function removeMonitoringTriggers_() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    const fn = t.getHandlerFunction();
    if (fn === 'runMonitoringWatchdog' || fn === 'sendDailyMonitoringSummary' || fn === 'rebuildMonitoringFleet') {
      ScriptApp.deleteTrigger(t);
    }
  });
}

// ============================================================
// WATCHDOG CHECKS
// ============================================================

function runMonitoringActiveHealthChecks_(ss, now) {
  checkGasApiHealth_(ss, now);
  checkGoogleSheetsDbHealth_(ss, now);
  checkAuthHealth_(ss, now);
  checkBackupHealth_(ss, now);
  checkMasterDataSyncHealth_(ss, now);
}

function checkGasApiHealth_(ss, now) {
  setHealthSnapshot_(
    ss,
    'GAS API',
    'OK',
    'GAS watchdog executed successfully',
    '',
    now
  );
}

function checkGoogleSheetsDbHealth_(ss, now) {
  try {
    const missing = [];

    Object.keys(MONITORING_TABS).forEach(function(name) {
      const sh = ss.getSheetByName(name);
      if (!sh) missing.push(name);
    });

    if (missing.length > 0) {
      setHealthSnapshot_(
        ss,
        'Google Sheets DB',
        'CRITICAL',
        'Missing monitoring sheets: ' + missing.join(', '),
        'Create missing monitoring sheets or fix sheet names',
        now
      );
      return;
    }

    setHealthSnapshot_(
      ss,
      'Google Sheets DB',
      'OK',
      'Monitoring spreadsheet and required tabs are accessible',
      '',
      now
    );

  } catch (err) {
    setHealthSnapshot_(
      ss,
      'Google Sheets DB',
      'CRITICAL',
      'Google Sheets DB check failed: ' + String(err && err.message ? err.message : err),
      'Check MONITORING_SPREADSHEET_ID and Apps Script permissions',
      now
    );
  }
}

function checkAuthHealth_(ss, now) {
  const secret = String(
    PropertiesService.getScriptProperties().getProperty(MONITORING_PROP_INGEST_SECRET) || ''
  ).trim();

  if (secret.length >= 16) {
    setHealthSnapshot_(
      ss,
      'Auth',
      'OK',
      'Monitoring ingest secret is configured',
      '',
      now
    );
    return;
  }

  setHealthSnapshot_(
    ss,
    'Auth',
    'CRITICAL',
    'Monitoring ingest secret is missing or too short',
    'Set MONITORING_INGEST_SECRET in Script Properties',
    now
  );
}

function checkBackupHealth_(ss, now) {
  const rows = readSheetAsObjects_(getMonitoringSheet_(ss, 'Backups'));

  if (!rows.length) {
    setHealthSnapshot_(
      ss,
      'Backup',
      'WARN',
      'No backup signal received yet',
      'Run backup test or wait for next scheduled backup',
      now
    );
    return;
  }

  let latestSuccess = null;
  let latestRow = null;

  rows.forEach(function(r) {
    const ts = parseDateSafe_(r.Timestamp);
    if (!ts) return;

    if (!latestRow || ts > latestRow.ts) {
      latestRow = { ts: ts, row: r };
    }

    if (String(r.Status || '').toUpperCase() === 'SUCCESS') {
      if (!latestSuccess || ts > latestSuccess) {
        latestSuccess = ts;
      }
    }
  });

  if (!latestSuccess) {
    setHealthSnapshot_(
      ss,
      'Backup',
      'CRITICAL',
      'No successful backup logged',
      'Run backup manually and verify backup procedure',
      now
    );
    return;
  }

  const ageHours = Math.floor((now.getTime() - latestSuccess.getTime()) / 3600000);

  if (ageHours > 48) {
    setHealthSnapshot_(
      ss,
      'Backup',
      'CRITICAL',
      'Last successful backup is older than 48h: ' + ageHours + 'h',
      'Run backup manually and verify backup procedure',
      now
    );
    return;
  }

  if (ageHours > 24) {
    setHealthSnapshot_(
      ss,
      'Backup',
      'WARN',
      'Last successful backup is older than 24h: ' + ageHours + 'h',
      'Check backup schedule',
      now
    );
    return;
  }

  setHealthSnapshot_(
    ss,
    'Backup',
    'OK',
    'Latest successful backup age: ' + ageHours + 'h',
    '',
    now
  );
}

function checkMasterDataSyncHealth_(ss, now) {
  const rows = readSheetAsObjects_(getMonitoringSheet_(ss, 'Events'));

  const successTypes = {
    MASTERDATA_SYNC_SUCCESS: true,
    MASTERDATA_SYNC_OK: true,
    STAMMDATEN_SYNC_SUCCESS: true,
    STAMMDATEN_SYNC_OK: true,
    GAS_MASTERDATA_SYNC_OK: true
  };

  const failTypes = {
    MASTERDATA_SYNC_FAIL: true,
    MASTERDATA_SYNC_ERROR: true,
    STAMMDATEN_SYNC_FAIL: true,
    STAMMDATEN_SYNC_ERROR: true,
    GAS_MASTERDATA_SYNC_FAIL: true
  };

  let latest = null;

  rows.forEach(function(r) {
    const eventType = String(r.EventType || '').trim();
    if (!successTypes[eventType] && !failTypes[eventType]) return;

    const ts = parseDateSafe_(r.Timestamp);
    if (!ts) return;

    if (!latest || ts > latest.ts) {
      latest = { ts: ts, row: r, eventType: eventType };
    }
  });

  if (!latest) {
    setHealthSnapshot_(
      ss,
      'MasterData Sync',
      'WARN',
      'No master data sync signal received yet',
      'Emit MASTERDATA_SYNC_SUCCESS or MASTERDATA_SYNC_FAIL from master data sync flow',
      now
    );
    return;
  }

  const ageHours = Math.floor((now.getTime() - latest.ts.getTime()) / 3600000);
  const msg = String(latest.row.Message || latest.eventType || '');

  if (failTypes[latest.eventType]) {
    setHealthSnapshot_(
      ss,
      'MasterData Sync',
      'CRITICAL',
      msg || 'Latest master data sync failed',
      'Check master data sync flow and source data',
      now
    );
    return;
  }

  if (ageHours > 24) {
    setHealthSnapshot_(
      ss,
      'MasterData Sync',
      'WARN',
      'Master data sync OK but older than 24h: ' + ageHours + 'h',
      'Run or verify master data sync schedule',
      now
    );
    return;
  }

  setHealthSnapshot_(
    ss,
    'MasterData Sync',
    'OK',
    msg || 'Master data sync OK',
    '',
    now
  );
}

function checkSEFStuck_(ss, now) {
  const sh = getMonitoringSheet_(ss, 'SEFStatus');
  const rows = readSheetAsObjects_(sh);
  rows.forEach(function(r) {
    const status = String(r.SEFStatus || '').toUpperCase();
    const nextAction = String(r.NextAction || '').toUpperCase();
    const ts = parseDateSafe_(r.LastAttemptAt || r.Timestamp);
    const ageMin = ts ? Math.floor((now.getTime() - ts.getTime()) / 60000) : 0;
    const corr = String(r.CorrelationId || r.InvoiceLocalId || r.BusinessInvoiceNo || '');

    if ((status === 'UNKNOWN' && ageMin > 30) || nextAction === 'MANUAL_REVIEW') {
      createMonitoringAlert_(ss, 'CRITICAL', 'SEF API', 'SEF_UNKNOWN_30_MIN', 'SEF UNKNOWN/manual review for ' + corr + ', ageMin=' + ageMin, corr);
    }

    if ((status === 'SENT' || status === 'SEND_STARTED' || status === 'STUCK') && ageMin > 60) {
      createMonitoringAlert_(ss, 'CRITICAL', 'SEF API', 'SEF_STUCK_60_MIN', 'SEF stuck for ' + corr + ', ageMin=' + ageMin, corr);
    }
  });
}

function checkSyncStuck_(ss, now) {
  const sh = getMonitoringSheet_(ss, 'SyncStatus');
  const rows = readSheetAsObjects_(sh);
  const latestByDevice = {};

  rows.forEach(function(r) {
    const deviceId = String(r.DeviceId || 'unknown');
    const ts = parseDateSafe_(r.Timestamp);
    if (!latestByDevice[deviceId] || (ts && ts > latestByDevice[deviceId].ts)) {
      latestByDevice[deviceId] = { row: r, ts: ts || new Date(0) };
    }
  });

  Object.keys(latestByDevice).forEach(function(deviceId) {
    const r = latestByDevice[deviceId].row;
    const pending = numberOrZero_(r.PendingCount);
    const failed = numberOrZero_(r.FailedCount);
    const conflicts = numberOrZero_(r.ConflictCount);
    const oldest = numberOrZero_(r.OldestPendingAgeMin);

    if (failed > 0) {
      createMonitoringAlert_(ss, 'WARN', 'PWA Sync', 'SYNC_FAILED_DEVICE', 'Device ' + deviceId + ' has failed sync items: ' + failed, deviceId);
    }

    if (oldest > 30 || conflicts > 0) {
      createMonitoringAlert_(ss, 'CRITICAL', 'Offline Queue', 'SYNC_PENDING_OR_CONFLICT', 'Device ' + deviceId + ' has old pending/conflicts. oldestMin=' + oldest + ', conflicts=' + conflicts + ', pending=' + pending, deviceId);
    }
  });
}

function checkBackupFreshness_(ss, now) {
  const sh = getMonitoringSheet_(ss, 'Backups');
  const rows = readSheetAsObjects_(sh);
  let latestSuccess = null;

  rows.forEach(function(r) {
    const status = String(r.Status || '').toUpperCase();
    const ts = parseDateSafe_(r.Timestamp);
    if (status === 'SUCCESS' && ts && (!latestSuccess || ts > latestSuccess)) {
      latestSuccess = ts;
    }
  });

  if (!latestSuccess) {
    createMonitoringAlert_(ss, 'CRITICAL', 'Backup', 'BACKUP_NEVER_SUCCESS', 'No successful backup logged', 'BACKUP');
    return;
  }

  const ageHours = Math.floor((now.getTime() - latestSuccess.getTime()) / 3600000);
  if (ageHours > 48) {
    createMonitoringAlert_(ss, 'CRITICAL', 'Backup', 'BACKUP_MISSING_48H', 'Last successful backup is older than 48h: ' + ageHours + 'h', 'BACKUP');
  } else if (ageHours > 24) {
    createMonitoringAlert_(ss, 'WARN', 'Backup', 'BACKUP_MISSING_24H', 'Last successful backup is older than 24h: ' + ageHours + 'h', 'BACKUP');
  }
}

function checkRecentGasErrors_(ss, now) {
  const sh = getMonitoringSheet_(ss, 'Errors');
  const rows = readSheetAsObjects_(sh);
  const cutoff = now.getTime() - 60 * 60 * 1000;
  let gasErrors = 0;

  rows.forEach(function(r) {
    const ts = parseDateSafe_(r.Timestamp);
    const source = String(r.Source || '').toUpperCase();
    if (ts && ts.getTime() >= cutoff && source === 'GAS') gasErrors++;
  });

  if (gasErrors >= 5) {
    createMonitoringAlert_(ss, 'WARN', 'GAS API', 'GAS_ERROR_SPIKE', 'GAS errors in last 60 min: ' + gasErrors, 'GAS');
  }
}

// ============================================================
// INIT / ENSURE WORKBOOK
// ============================================================

function initializeMonitoringWorkbook() {
  return monitoringWithLock_(function() {
    const ss = getMonitoringSpreadsheet_();
    ensureMonitoringWorkbook_(ss);
    ensureDefaultHealthRows_(ss);
    return {
      success: true,
      spreadsheetId: ss.getId(),
      spreadsheetName: ss.getName()
    };
  });
}

function ensureMonitoringWorkbook_(ss) {
  Object.keys(MONITORING_TABS).forEach(function(name) {
    const sh = getMonitoringSheet_(ss, name);
    ensureHeader_(sh, MONITORING_TABS[name]);
  });
  ensureDefaultHealthRows_(ss);
}

function ensureDefaultHealthRows_(ss) {
  const sh = getMonitoringSheet_(ss, 'Health');
  MONITORING_COMPONENTS.forEach(function(component) {
    if (findRowByFirstColumn_(sh, component) < 0) {
      sh.appendRow([
        component,
        'UNKNOWN',
        '',
        '',
        0,
        0,
        0,
        'No events yet',
        'No action until first event/watchdog check',
        new Date().toISOString()
      ]);
    }
  });
}

function getMonitoringSpreadsheet_() {
  const props = PropertiesService.getScriptProperties();

  // 1) Eksplicitni ID iz Script Properties (kanonski kljuc, uz legacy fallback).
  let storedId = String(props.getProperty(MONITORING_PROP_SPREADSHEET_ID) || '').trim();
  if (!storedId) {
    storedId = String(props.getProperty(MONITORING_PROP_SPREADSHEET_ID_LEGACY) || '').trim();
  }
  if (storedId) {
    try {
      return SpreadsheetApp.openById(storedId);
    } catch (errOpen) {
      // ID vise nije upotrebljiv (obrisan / u trash-u / nema pristup) -> nastavi na discovery/create.
      try { Logger.log('Stored monitoring spreadsheet id not usable: ' + errOpen.message); } catch (ignore) {}
    }
  }

  // 2) Postojeci fajl po kanonskom imenu -> zapamti ID za ubuduce.
  const files = DriveApp.getFilesByName(MONITORING_SPREADSHEET_NAME);
  if (files.hasNext()) {
    const ss = SpreadsheetApp.open(files.next());
    rememberMonitoringSpreadsheetId_(props, ss.getId());
    return ss;
  }

  // 3) Ne postoji nista -> auto-kreiraj kanonski workbook i zapamti ID.
  //    Tako nema potrebe za rucnim setovanjem Script Property na svakoj instalaciji.
  const created = SpreadsheetApp.create(MONITORING_SPREADSHEET_NAME);
  rememberMonitoringSpreadsheetId_(props, created.getId());
  return created;
}

function rememberMonitoringSpreadsheetId_(props, id) {
  if (!id) return;
  try {
    props.setProperty(MONITORING_PROP_SPREADSHEET_ID, id);
  } catch (err) {
    try { Logger.log('Could not persist ' + MONITORING_PROP_SPREADSHEET_ID + ': ' + (err && err.message ? err.message : err)); } catch (ignore) {}
  }
}

function getMonitoringSheet_(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function ensureHeader_(sh, headers) {
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sh.setFrozenRows(1);
    return;
  }

  const existing = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), headers.length)).getValues()[0];
  let mismatch = false;
  for (let i = 0; i < headers.length; i++) {
    if (String(existing[i] || '') !== headers[i]) {
      mismatch = true;
      break;
    }
  }

  if (mismatch) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
}

// ============================================================
// HELPERS
// ============================================================

function monitoringWithLock_(fn) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    return fn();
  } catch (err) {
    try {
      Logger.log('monitoringWithLock failed: ' + err.message);
    } catch (ignore) {}
    return { success: false, error: err.message || String(err) };
  } finally {
    try { lock.releaseLock(); } catch (ignore2) {}
  }
}

function isValidMonitoringSecret_(secret) {
  const expected = String(PropertiesService.getScriptProperties().getProperty(MONITORING_PROP_INGEST_SECRET) || '').trim();
  if (!expected) return false;
  return String(secret || '') === expected;
}

function getMonitoringAlertEmail_() {
  return String(PropertiesService.getScriptProperties().getProperty(MONITORING_PROP_ALERT_EMAIL) || '').trim();
}

function normalizeSource_(source) {
  const s = String(source || '').toUpperCase();
  if (s === 'PWA') return 'PWA';
  if (s === 'VBA') return 'VBA';
  if (s === 'SEF') return 'SEF';
  if (s === 'BACKUP') return 'BACKUP';
  if (s === 'SYSTEM') return 'SYSTEM';
  return 'GAS';
}

function normalizeSeverity_(severity) {
  const s = String(severity || '').toUpperCase();
  if (s === 'CRITICAL') return 'CRITICAL';
  if (s === 'ERROR') return 'ERROR';
  if (s === 'WARN' || s === 'WARNING') return 'WARN';
  return 'INFO';
}

function inferComponent_(source, eventType, explicitComponent) {
  if (explicitComponent) return String(explicitComponent);
  const e = String(eventType || '').toUpperCase();
  const s = String(source || '').toUpperCase();

  if (s === 'SEF' || e.indexOf('SEF_') === 0) return 'SEF API';
  if (s === 'VBA' || e.indexOf('VBA_') === 0) return 'VBA Client';
  if (s === 'BACKUP' || e.indexOf('BACKUP_') === 0) return 'Backup';
  if (e.indexOf('SYNC_') === 0 || e === 'PWA_HEARTBEAT') return 'PWA Sync';
  if (e.indexOf('AUTH_') === 0 || e.indexOf('LOGIN_') === 0) return 'Auth';
  if (e.indexOf('MASTERDATA_') === 0 || e.indexOf('STAMMDATEN_') === 0) return 'MasterData Sync';
  return 'GAS API';
}

function statusFromSeverity_(severity) {
  const s = normalizeSeverity_(severity);
  if (s === 'CRITICAL') return 'CRITICAL';
  if (s === 'ERROR' || s === 'WARN') return 'WARN';
  return 'OK';
}

function worseStatus_(a, b) {
  const rank = { UNKNOWN: 0, OK: 1, WARN: 2, CRITICAL: 3 };
  const aa = String(a || 'UNKNOWN').toUpperCase();
  const bb = String(b || 'UNKNOWN').toUpperCase();
  return (rank[bb] || 0) > (rank[aa] || 0) ? bb : aa;
}

function worseSeverity_(a, b) {
  const rank = { INFO: 1, WARN: 2, ERROR: 3, CRITICAL: 4 };
  const aa = normalizeSeverity_(a);
  const bb = normalizeSeverity_(b);
  return (rank[bb] || 0) > (rank[aa] || 0) ? bb : aa;
}

function ownerActionForStatus_(component, status, event) {
  if (status === 'OK') return '';
  if (component === 'SEF API') return 'Check SEFStatus row and verify in SEF portal before retry';
  if (component === 'PWA Sync' || component === 'Offline Queue') return 'Check SyncStatus for device/user and inspect failed queue';
  if (component === 'Backup') return 'Run backup manually and verify restore path';
  if (component === 'VBA Client') return 'Check VBA module/procedure and affected entity';
  return 'Inspect Errors and Events by CorrelationId';
}

function isSyncEvent_(event) {
  const e = String(event.eventType || '').toUpperCase();
  return e === 'PWA_HEARTBEAT' || e.indexOf('SYNC_') === 0;
}

function isSessionEvent_(event) {
  const e = String(event.eventType || '').toUpperCase();
  return e === 'PWA_HEARTBEAT' || e === 'APP_OPEN' || e === 'LOGIN_SUCCESS' || e === 'VBA_APP_OPEN';
}

function isSEFEvent_(event) {
  const e = String(event.eventType || '').toUpperCase();
  return event.source === 'SEF' || e.indexOf('SEF_') === 0;
}

function isBackupEvent_(event) {
  const e = String(event.eventType || '').toUpperCase();
  return event.source === 'BACKUP' || e.indexOf('BACKUP_') === 0;
}

function isCriticalAuditEvent_(event) {
  const e = String(event.eventType || '').toUpperCase();
  return event.severity === 'CRITICAL' ||
    e.indexOf('PAYMENT_') === 0 ||
    e.indexOf('FAKTURA_') === 0 ||
    e.indexOf('SEF_') === 0 ||
    e.indexOf('MANUAL_OVERRIDE') >= 0;
}

function inferSEFStatusFromEvent_(event) {
  const e = String(event.eventType || '').toUpperCase();
  if (e === 'SEF_SEND_START') return 'SEND_STARTED';
  if (e === 'SEF_SEND_SUCCESS') return 'SENT';
  if (e === 'SEF_SEND_FAIL') return 'FAILED';
  if (e === 'SEF_STATUS_ACCEPTED') return 'ACCEPTED';
  if (e === 'SEF_STATUS_REJECTED') return 'REJECTED';
  if (e === 'SEF_UNKNOWN') return 'UNKNOWN';
  return '';
}

function inferSEFNextAction_(event) {
  if (event.severity === 'CRITICAL') return 'MANUAL_REVIEW';
  const e = String(event.eventType || '').toUpperCase();
  if (e === 'SEF_SEND_FAIL') return 'RETRY';
  if (e === 'SEF_UNKNOWN') return 'CHECK_SEF_PORTAL';
  return 'WAIT';
}

function inferBackupStatusFromEvent_(event) {
  const e = String(event.eventType || '').toUpperCase();
  if (e === 'BACKUP_SUCCESS') return 'SUCCESS';
  if (e === 'BACKUP_FAIL') return 'FAILED';
  return event.severity === 'ERROR' || event.severity === 'CRITICAL' ? 'FAILED' : 'SUCCESS';
}

function findSEFStatusRow_(sh, correlationId, invoiceLocalId) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return -1;
  const values = sh.getRange(2, 1, lastRow - 1, Math.max(sh.getLastColumn(), 14)).getValues();
  for (let i = 0; i < values.length; i++) {
    const rowInvoiceLocalId = String(values[i][1] || '');
    const rowCorrelationId = String(values[i][12] || '');
    if (correlationId && rowCorrelationId === correlationId) return i + 2;
    if (invoiceLocalId && rowInvoiceLocalId === invoiceLocalId) return i + 2;
  }
  return -1;
}

function findRowByFirstColumn_(sh, value) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return -1;
  const values = sh.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || '') === String(value || '')) return i + 2;
  }
  return -1;
}

function getExistingCell_(sh, rowIndex, colIndex) {
  if (rowIndex < 2) return '';
  return sh.getRange(rowIndex, colIndex).getValue();
}

function openAlertExists_(sh, component, alertCode, correlationId) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return false;
  const values = sh.getRange(2, 1, lastRow - 1, Math.max(sh.getLastColumn(), 12)).getValues();
  for (let i = 0; i < values.length; i++) {
    const rowComponent = String(values[i][2] || '');
    const rowAlertCode = String(values[i][3] || '');
    const rowCorrelationId = String(values[i][5] || '');
    const resolved = values[i][10] === true || String(values[i][10]).toLowerCase() === 'true';
    if (!resolved && rowComponent === component && rowAlertCode === alertCode && rowCorrelationId === String(correlationId || '')) {
      return true;
    }
  }
  return false;
}

function readOpenAlerts_(ss) {
  const sh = getMonitoringSheet_(ss, 'Alerts');
  const all = readSheetAsObjects_(sh);
  return all.filter(function(r) {
    return String(r.Resolved || '').toLowerCase() !== 'true';
  });
}

function readSheetAsObjects_(sh) {
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];
  const values = sh.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = values[0].map(function(h) { return String(h || ''); });
  const out = [];
  for (let r = 1; r < values.length; r++) {
    const obj = {};
    for (let c = 0; c < headers.length; c++) {
      if (headers[c]) obj[headers[c]] = values[r][c];
    }
    obj._rowIndex = r + 1;
    out.push(obj);
  }
  return out;
}

function countRecentEventsByComponent_(ss, component, hours) {
  const sh = getMonitoringSheet_(ss, 'Events');
  const rows = readSheetAsObjects_(sh);
  const cutoff = Date.now() - hours * 3600000;
  const counts = { ERROR: 0, WARN: 0, CRITICAL: 0 };

  rows.forEach(function(r) {
    const ts = parseDateSafe_(r.Timestamp);
    if (!ts || ts.getTime() < cutoff) return;
    const inferredComponent = inferComponent_(r.Source, r.EventType, '');
    if (inferredComponent !== component) return;
    const sev = normalizeSeverity_(r.Severity);
    if (counts[sev] !== undefined) counts[sev]++;
  });

  return counts;
}

function sanitizePayload_(payload) {
  let clone = {};
  try {
    clone = JSON.parse(JSON.stringify(payload || {}));
  } catch (err) {
    clone = { value: String(payload || '') };
  }

  return sanitizeObjectRecursive_(clone);
}

function sanitizeObjectRecursive_(obj) {
  if (obj === null || obj === undefined) return obj;
  if (typeof obj !== 'object') return obj;

  if (Array.isArray(obj)) {
    return obj.slice(0, 50).map(function(x) { return sanitizeObjectRecursive_(x); });
  }

  const blocked = [
    'token',
    'password',
    'pin',
    'refreshToken',
    'accessToken',
    'apiKey',
    'sefApiKey',
    'authorization',
    'xml',
    'pdf',
    'base64',
    'fileContent'
  ];

  const out = {};
  Object.keys(obj).forEach(function(k) {
    const lower = String(k).toLowerCase();
    const blockedHit = blocked.some(function(b) { return lower.indexOf(b.toLowerCase()) >= 0; });
    if (blockedHit) {
      out[k] = '[REDACTED]';
    } else {
      out[k] = sanitizeObjectRecursive_(obj[k]);
    }
  });
  return out;
}

function stringifyPayload_(payload) {
  let s = '';
  try {
    s = JSON.stringify(payload || {});
  } catch (err) {
    s = String(payload || '');
  }
  return truncate_(s, MONITORING_MAX_PAYLOAD_CHARS);
}

function truncate_(value, maxLen) {
  const s = String(value || '');
  if (s.length <= maxLen) return s;
  return s.substring(0, maxLen) + '...';
}

function numberOrZero_(v) {
  const n = Number(v);
  return isNaN(n) ? 0 : n;
}

function boolToText_(v) {
  if (v === true || String(v).toLowerCase() === 'true' || String(v).toLowerCase() === 'online') return 'TRUE';
  if (v === false || String(v).toLowerCase() === 'false' || String(v).toLowerCase() === 'offline') return 'FALSE';
  return String(v || '');
}

function parseDateSafe_(v) {
  if (!v) return null;
  const d = new Date(v);
  if (isNaN(d.getTime())) return null;
  return d;
}

function componentKey_(component) {
  return String(component || '').toUpperCase().replace(/[^A-Z0-9]+/g, '_').replace(/^_+|_+$/g, '');
}

function setHealthSnapshot_(ss, component, status, message, ownerAction, now) {
  const sh = getMonitoringSheet_(ss, 'Health');
  const idx = findRowByFirstColumn_(sh, component);
  const counts = countRecentEventsByComponent_(ss, component, 24);
  const timestamp = now.toISOString();

  const normalizedStatus = String(status || 'UNKNOWN').toUpperCase();
  const isOk = normalizedStatus === 'OK';

  const row = [
    component,
    normalizedStatus,
    isOk ? timestamp : getExistingCell_(sh, idx, 3),
    !isOk ? timestamp : getExistingCell_(sh, idx, 4),
    counts.ERROR,
    counts.WARN,
    counts.CRITICAL,
    message || normalizedStatus,
    ownerAction || '',
    timestamp
  ];

  if (idx > 0) {
    sh.getRange(idx, 1, 1, row.length).setValues([row]);
  } else {
    sh.appendRow(row);
  }
}
// ============================================================
// OPTIONAL TEST HELPERS
// ============================================================

function testMonitoringIngestPWA() {
  return monitoringIngest_({
    action: 'monitor',
    source: 'PWA',
    eventType: 'PWA_HEARTBEAT',
    severity: 'INFO',
    userId: 'test-user',
    role: 'Otkupac',
    deviceId: 'test-device',
    appVersion: 'test-1.0.0',
    payload: {
      online: true,
      route: '/otkup/form',
      pendingCount: 0,
      failedCount: 0,
      conflictCount: 0,
      oldestPendingAgeMin: 0,
      lastSyncAt: new Date().toISOString()
    }
  }, { entityID: 'test-user', role: 'Otkupac' });
}

function testMonitoringIngestSEFUnknown() {
  return monitoringIngest_({
    action: 'monitor',
    source: 'SEF',
    eventType: 'SEF_UNKNOWN',
    severity: 'CRITICAL',
    entityType: 'Faktura',
    entityId: 'FAK-TEST-001',
    correlationId: 'FAK-TEST-001',
    message: 'Timeout after SEF send; status unknown',
    payload: {
      invoiceLocalId: 'FAK-TEST-001',
      businessInvoiceNo: 'TEST-001/2026',
      sefRequestId: 'REQ-TEST-001',
      sefStatus: 'UNKNOWN',
      localStatus: 'SENDING',
      attemptCount: 1,
      lastHttpCode: 'timeout',
      nextAction: 'CHECK_SEF_PORTAL',
      needsManualReview: true
    }
  }, { entityID: 'operator', role: 'Management' });
}
