const AGRIX_FOLDER_PROPS = {
  ROOT: 'AGRIX_ROOT_FOLDER_ID',

  INBOX: 'AGRIX_INBOX_FOLDER_ID',

  SHEETS: 'AGRIX_SHEETS_FOLDER_ID',
  SHEETS_OPERATIONAL: 'AGRIX_SHEETS_OPERATIONAL_FOLDER_ID',
  SHEETS_MASTER: 'AGRIX_SHEETS_MASTER_FOLDER_ID',
  SHEETS_REPORTS: 'AGRIX_SHEETS_REPORTS_FOLDER_ID',
  SHEETS_ARCHIVE: 'AGRIX_SHEETS_ARCHIVE_FOLDER_ID',

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

function getMgmtReportsSpreadsheet_() {
  return getSpreadsheetByNameInFolder_('SHEETS_REPORTS', 'MgmtReports');
}

function getKarticeSpreadsheet_() {
  return getSpreadsheetByNameInFolder_('SHEETS_REPORTS', 'Kartice');
}

// ============================================================
// One-time bootstrap za novog klijenta (C002, C003, ...).
// Napravi celo Drive stablo ispod root foldera I upiše svaki
// folder ID u Script Properties. Time se sažimaju koraci §5
// (stablo), §8 (skupljanje ID-jeva) i §11 (upis u Script
// Properties) iz install/AgriX_Onboarding_Vodic_Novi_Klijent_v2.md.
//
// VAŽNO: pokreni ovo UNUTAR GAS projekta TOG klijenta
// (AgriX_C00X_GAS_PROD) — Script Properties su per-projekat.
// Idempotentno je: postojeći folder se po imenu reuse-uje, ne dupliraju se.
// ============================================================

// Stablo foldera (vodič §5). `prop` ključevi odgovaraju AGRIX_FOLDER_PROPS gore.
// Čvorovi bez `prop` se prave ali se ne upisuju u Script Properties (organizacioni).
const AGRIX_FOLDER_TREE = [
  { name: '00_Inbox', prop: 'INBOX', children: [
    // '01_Bank' = odrediste Bank PDF Gmail Downloader-a (driveFolderId u
    // BANK_IMPORT_CLIENTS_JSON) i BANKA_DRIVE_SOURCE_PATH koji VBA puller cita.
    // 'Downloaded' = sibling u koji puller premesta vec povucene PDF-ove
    // (default BANKA_DRIVE_DOWNLOADED_PATH = <parent(source)>\Downloaded).
    { name: '01_Bank' }, { name: 'Downloaded' },
    { name: 'Fiskalni' }, { name: 'Uvoz' }, { name: 'Manual' }
  ]},
  { name: '01_Sheets', prop: 'SHEETS', children: [
    { name: '01_Operational', prop: 'SHEETS_OPERATIONAL' },
    { name: '02_Master', prop: 'SHEETS_MASTER' },
    { name: '03_Reports', prop: 'SHEETS_REPORTS' },
    { name: '04_Archive', prop: 'SHEETS_ARCHIVE' }
  ]},
  { name: '03_Documents', prop: 'DOCUMENTS', children: [
    { name: 'Otkupni_Listovi', prop: 'DOC_OTKUPNI_LISTOVI' },
    { name: 'Otpremnice', prop: 'DOC_OTPREMNICE' },
    { name: 'Zbirne', prop: 'DOC_ZBIRNE' },
    { name: 'Fakture', prop: 'DOC_FAKTURE' }
  ]},
  { name: '04_Export', prop: 'EXPORT', children: [
    { name: 'Excel', prop: 'EXPORT_EXCEL' },
    { name: 'PDF', prop: 'EXPORT_PDF' },
    { name: 'CSV', prop: 'EXPORT_CSV' },
    { name: 'API', prop: 'EXPORT_API' }
  ]},
  { name: '05_Backup', prop: 'BACKUP', children: [
    { name: 'Daily', prop: 'BACKUP_DAILY' },
    { name: 'Weekly', prop: 'BACKUP_WEEKLY' },
    { name: 'Before_Release', prop: 'BACKUP_BEFORE_RELEASE' }
  ]},
  { name: '06_Monitoring', prop: 'MONITORING', children: [
    { name: 'ErrorLog', prop: 'MONITORING_ERRORLOG' },
    { name: 'Sync_Reports', prop: 'MONITORING_SYNC_REPORTS' },
    { name: 'Incidenti', prop: 'MONITORING_INCIDENTI' },
    { name: 'Health_Checks', prop: 'MONITORING_HEALTH_CHECKS' }
  ]},
  { name: '07_Admin', prop: 'ADMIN', children: [
    { name: 'Config', prop: 'ADMIN_CONFIG' },
    { name: 'Deployments', prop: 'ADMIN_DEPLOYMENTS' },
    { name: 'Access', prop: 'ADMIN_ACCESS' },
    { name: 'Templates', prop: 'ADMIN_TEMPLATES' },
    { name: 'Runbooks', prop: 'ADMIN_RUNBOOKS' }
  ]}
];

function getOrCreateChildFolder_(parentFolder, name) {
  const existing = parentFolder.getFoldersByName(name);
  if (existing.hasNext()) {
    return existing.next();
  }
  return parentFolder.createFolder(name);
}

// Rekurzivno pravi stablo ispod `parentFolder`; za svaki čvor sa `prop`
// upisuje {propName: folderId} u `propMap`, a u `report` ide čitljiv prikaz.
function buildAgriXTree_(parentFolder, nodes, propMap, report, indent) {
  nodes.forEach(function (node) {
    const folder = getOrCreateChildFolder_(parentFolder, node.name);
    report.push(indent + node.name + '  ->  ' + folder.getId());

    if (node.prop) {
      const propName = AGRIX_FOLDER_PROPS[node.prop];
      if (!propName) {
        throw new Error('Tree references unknown folder key: ' + node.prop);
      }
      propMap[propName] = folder.getId();
    }

    if (node.children && node.children.length) {
      buildAgriXTree_(folder, node.children, propMap, report, indent + '  ');
    }
  });
}

// MAIN — pokreni jednom po klijentu, u GAS projektu tog klijenta.
// Koraci:
//   1. Ručno napravi root folder AgriX_C00X_PROD i podeli ga sa backup@ (vodič §5–6).
//   2. Ulepi njegov folder ID u ROOT_FOLDER_ID ispod.
//   3. Run > bootstrapAgriXFolderTree  → napravi sve podfoldere + upiše Script Properties.
//   4. Potom pokreni debugAgriXFolders (vodič §13) — svi moraju biti OK.
function bootstrapAgriXFolderTree() {
  const ROOT_FOLDER_ID = 'PASTE_AgriX_C00X_PROD_FOLDER_ID_HERE';

  if (!ROOT_FOLDER_ID || ROOT_FOLDER_ID.indexOf('PASTE_') === 0) {
    throw new Error('Postavi ROOT_FOLDER_ID na ID foldera AgriX_C00X_PROD pre pokretanja.');
  }

  const root = DriveApp.getFolderById(ROOT_FOLDER_ID);
  const propMap = {};
  const report = [];

  // Upiši i sam root.
  propMap[AGRIX_FOLDER_PROPS.ROOT] = root.getId();
  report.push(root.getName() + '  (ROOT)  ->  ' + root.getId());

  buildAgriXTree_(root, AGRIX_FOLDER_TREE, propMap, report, '  ');

  // Upiši sve sakupljene ID-jeve u Script Properties u jednom batch-u.
  // deleteAllOthers = false → ne briše postojeće propse (npr. MONITORING_SPREADSHEET_ID, tajne).
  PropertiesService.getScriptProperties().setProperties(propMap, false);

  report.push('');
  report.push('Script Properties upisano: ' + Object.keys(propMap).length + ' (očekivano 33).');
  Logger.log(report.join('\n'));
  return report;
}
