const AGRIX_FOLDER_PROPS = {
  ROOT: 'AGRIX_ROOT_FOLDER_ID',

  INBOX: 'AGRIX_INBOX_FOLDER_ID',

  SHEETS: 'AGRIX_SHEETS_FOLDER_ID',
  SHEETS_OPERATIONAL: 'AGRIX_SHEETS_OPERATIONAL_FOLDER_ID',
  SHEETS_MASTER: 'AGRIX_SHEETS_MASTER_FOLDER_ID',
  SHEETS_REPORTS: 'AGRIX_SHEETS_REPORTS_FOLDER_ID',
  SHEETS_ARCHIVE: 'AGRIX_SHEETS_ARCHIVE_FOLDER_ID',

  BANK_IZVODI: 'AGRIX_BANK_IZVODI_FOLDER_ID',

  DOCUMENTS: 'AGRIX_DOCUMENTS_FOLDER_ID',
  DOC_OTKUPNI_LISTOVI: 'AGRIX_DOC_OTKUPNI_LISTOVI_FOLDER_ID',
  DOC_OTPREMNICE: 'AGRIX_DOC_OTPREMNICE_FOLDER_ID',
  DOC_ZBIRNE: 'AGRIX_DOC_ZBIRNE_FOLDER_ID',
  DOC_FAKTURE: 'AGRIX_DOC_FAKTURE_FOLDER_ID',

  EXPORT: 'AGRIX_EXPORT_FOLDER_ID',
  EXPORT_EXCEL: 'AGRIX_EXPORT_EXCEL_FOLDER_ID',
  EXPORT_PDF: 'AGRIX_EXPORT_PDF_FOLDER_ID',
  EXPORT_CSV: 'AGRIX_EXPORT_CSV_FOLDER_ID',
  EXPORT_API: 'AGRIX_EXPORT_API_FOLDER_ID',

  BACKUP: 'AGRIX_BACKUP_FOLDER_ID',
  BACKUP_DAILY: 'AGRIX_BACKUP_DAILY_FOLDER_ID',
  BACKUP_WEEKLY: 'AGRIX_BACKUP_WEEKLY_FOLDER_ID',
  BACKUP_BEFORE_RELEASE: 'AGRIX_BACKUP_BEFORE_RELEASE_FOLDER_ID',

  MONITORING: 'AGRIX_MONITORING_FOLDER_ID',
  MONITORING_ERRORLOG: 'AGRIX_MONITORING_ERRORLOG_FOLDER_ID',
  MONITORING_SYNC_REPORTS: 'AGRIX_MONITORING_SYNC_REPORTS_FOLDER_ID',
  MONITORING_INCIDENTI: 'AGRIX_MONITORING_INCIDENTI_FOLDER_ID',
  MONITORING_HEALTH_CHECKS: 'AGRIX_MONITORING_HEALTH_CHECKS_FOLDER_ID',

  ADMIN: 'AGRIX_ADMIN_FOLDER_ID',
  ADMIN_CONFIG: 'AGRIX_ADMIN_CONFIG_FOLDER_ID',
  ADMIN_DEPLOYMENTS: 'AGRIX_ADMIN_DEPLOYMENTS_FOLDER_ID',
  ADMIN_ACCESS: 'AGRIX_ADMIN_ACCESS_FOLDER_ID',
  ADMIN_TEMPLATES: 'AGRIX_ADMIN_TEMPLATES_FOLDER_ID',
  ADMIN_RUNBOOKS: 'AGRIX_ADMIN_RUNBOOKS_FOLDER_ID'
};

function getAgriXFolder_(key) {
  const propName = AGRIX_FOLDER_PROPS[key];
  if (!propName) {
    throw new Error('Unknown AgriX folder key: ' + key);
  }

  const folderId = PropertiesService
    .getScriptProperties()
    .getProperty(propName);

  if (!folderId) {
    throw new Error('Missing Script Property: ' + propName);
  }

  return DriveApp.getFolderById(folderId);
}

function getSpreadsheetByNameInFolder_(folderKey, spreadsheetName) {
  const folder = getAgriXFolder_(folderKey);
  const files = folder.getFilesByName(spreadsheetName);

  if (!files.hasNext()) {
    return null;
  }

  return SpreadsheetApp.open(files.next());
}

function createSpreadsheetInFolder_(folderKey, spreadsheetName, headerRow) {
  const folder = getAgriXFolder_(folderKey);
  const ss = SpreadsheetApp.create(spreadsheetName);
  const file = DriveApp.getFileById(ss.getId());

  file.moveTo(folder);

  if (headerRow && headerRow.length) {
    const sh = ss.getSheets()[0];
    sh.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);
    sh.getRange(1, 1, 1, headerRow.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }

  return ss;
}

function getOrCreateSpreadsheetInFolder_(folderKey, spreadsheetName, headerRow) {
  let ss = getSpreadsheetByNameInFolder_(folderKey, spreadsheetName);

  if (ss) {
    return ss;
  }

  return createSpreadsheetInFolder_(folderKey, spreadsheetName, headerRow || []);
}

function getStammdatenSpreadsheet_() {
  const ss = getSpreadsheetByNameInFolder_('SHEETS_MASTER', 'Stammdaten');

  if (!ss) {
    throw new Error('Stammdaten spreadsheet not found in AGRIX_SHEETS_MASTER_FOLDER_ID');
  }

  return ss;
}

function getOperationalSpreadsheet_(spreadsheetName, headerRow) {
  return getOrCreateSpreadsheetInFolder_(
    'SHEETS_OPERATIONAL',
    spreadsheetName,
    headerRow || []
  );
}

function getReportSpreadsheet_(spreadsheetName) {
  const ss = getSpreadsheetByNameInFolder_('SHEETS_REPORTS', spreadsheetName);

  if (!ss) {
    throw new Error('Report spreadsheet not found in AGRIX_SHEETS_REPORTS_FOLDER_ID: ' + spreadsheetName);
  }

  return ss;
}

function getMonitoringErrorLogSpreadsheet_() {
  return getOrCreateSpreadsheetInFolder_(
    'MONITORING_ERRORLOG',
    'ErrorLog',
    ['Timestamp', 'Source', 'Action', 'Message', 'Details', 'EntityID', 'Severity']
  );
}
