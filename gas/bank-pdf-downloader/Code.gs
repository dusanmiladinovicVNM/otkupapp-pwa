/*******************************************************
 * Multi-client Bank PDF Gmail Downloader
 *
 * Install on the REAL Gmail account that receives bank emails.
 *
 * Responsibility:
 * Gmail PDF attachments -> shared Drive folder(s)
 *
 * This script DOES NOT:
 * - parse bank PDFs
 * - write to Excel/VBA tables
 * - write to tblBankaImport
 * - map to tblNovac
 *
 * Existing OtkupApp/VBA BankaImport remains the owner of:
 * PDF parsing, saldo integrity, staging and mapping.
 *******************************************************/


/* ======================================================
 * Property names
 * ====================================================== */

const BANK_PROP_CLIENTS_JSON = "BANK_IMPORT_CLIENTS_JSON";
const BANK_PROP_SECRET = "BANK_IMPORT_SECRET";
const BANK_PROP_BACKFILL_FROM_DATE = "BANK_IMPORT_BACKFILL_FROM_DATE";
const BANK_PROP_BACKFILL_TO_DATE = "BANK_IMPORT_BACKFILL_TO_DATE";


/* ======================================================
 * Public entrypoints
 * ====================================================== */

/**
 * Daily / normal import.
 * Processes all enabled clients from BANK_IMPORT_CLIENTS_JSON.
 */
function runBankPdfImportDaily() {
  return runBankPdfImport_("daily", {});
}


/**
 * Manual run from Apps Script editor.
 * Same as daily, useful for testing/operator refresh.
 */
function runBankPdfImportNow() {
  return runBankPdfImport_("manual", {});
}


/**
 * Backfill run.
 *
 * Uses Script Properties:
 * BANK_IMPORT_BACKFILL_FROM_DATE = YYYY-MM-DD
 * BANK_IMPORT_BACKFILL_TO_DATE   = YYYY-MM-DD, optional
 *
 * If TO date is empty, today is used.
 */
function runBankPdfImportBackfill() {
  const props = PropertiesService.getScriptProperties();

  const fromDate = String(
    props.getProperty(BANK_PROP_BACKFILL_FROM_DATE) || ""
  ).trim();

  const toDate = String(
    props.getProperty(BANK_PROP_BACKFILL_TO_DATE) || ""
  ).trim();

  if (!fromDate) {
    throw new Error(
      "Missing Script Property " +
      BANK_PROP_BACKFILL_FROM_DATE +
      " in YYYY-MM-DD format."
    );
  }

  return runBankPdfImport_("backfill", {
    fromDate: fromDate,
    toDate: toDate || getTodayIsoDate_()
  });
}


/**
 * Config + Gmail + Drive test.
 * Does not save files.
 */
function testBankPdfImportConfig() {
  const clients = getEnabledClients_();
  const result = {
    success: true,
    clientCount: clients.length,
    clients: []
  };

  for (let i = 0; i < clients.length; i++) {
    const client = clients[i];
    const folder = DriveApp.getFolderById(client.driveFolderId);
    const query = buildGmailQuery_(client, "test", {});

    const sampleThreads = GmailApp.search(query, 0, 5);

    result.clients.push({
      clientId: client.clientId,
      folderName: folder.getName(),
      query: query,
      sampleThreadCount: sampleThreads.length
    });
  }

  console.log(JSON.stringify(result, null, 2));
  return result;
}


/**
 * Gmail-only test.
 * If this fails with "Gmail operation not allowed",
 * the executing account is not a real Gmail mailbox or GmailApp is blocked.
 */
function testGmailAccessOnly() {
  const threads = GmailApp.search("in:inbox newer_than:30d", 0, 1);

  const result = {
    success: true,
    threadCount: threads.length
  };

  console.log(JSON.stringify(result, null, 2));
  return result;
}


/**
 * Installs one daily trigger.
 * Run once manually after config test passes.
 */
function setupDailyBankPdfImportTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === "runBankPdfImportDaily") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger("runBankPdfImportDaily")
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .create();

  console.log("Daily trigger installed for runBankPdfImportDaily.");
}


/**
 * Removes the daily trigger.
 */
function removeDailyBankPdfImportTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === "runBankPdfImportDaily") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  console.log("Daily trigger removed.");
}


/**
 * Optional WebApp endpoint for VBA/manual refresh.
 *
 * Deploy:
 * - Execute as: Me
 * - Who has access: Anyone with the link
 *
 * Required JSON:
 * {
 *   "action": "runNow",
 *   "secret": "..."
 * }
 *
 * Optional:
 * {
 *   "action": "backfill",
 *   "secret": "...",
 *   "fromDate": "2026-01-01",
 *   "toDate": "2026-07-01",
 *   "clientId": "CLIENT_A"
 * }
 */
function doPost(e) {
  try {
    const data = parsePostJson_(e);
    validateWebAppSecret_(data);

    const action = String(data.action || "").trim();

    if (action === "runNow") {
      return jsonResponse_(
        runBankPdfImport_("webapp-manual", {
          clientId: String(data.clientId || "").trim()
        })
      );
    }

    if (action === "backfill") {
      const fromDate = String(data.fromDate || "").trim();
      const toDate = String(data.toDate || "").trim() || getTodayIsoDate_();

      if (!fromDate) {
        return jsonResponse_({
          success: false,
          code: "MISSING_FROM_DATE",
          message: "fromDate is required for backfill."
        });
      }

      return jsonResponse_(
        runBankPdfImport_("webapp-backfill", {
          fromDate: fromDate,
          toDate: toDate,
          clientId: String(data.clientId || "").trim()
        })
      );
    }

    if (action === "test") {
      return jsonResponse_(testBankPdfImportConfig());
    }

    return jsonResponse_({
      success: false,
      code: "UNKNOWN_ACTION",
      message: "Unknown action: " + action
    });

  } catch (err) {
    return jsonResponse_({
      success: false,
      code: "ERROR",
      message: err && err.message ? err.message : String(err)
    });
  }
}


/* ======================================================
 * Main runner
 * ====================================================== */

function runBankPdfImport_(mode, options) {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(10000)) {
    return {
      success: false,
      code: "LOCKED",
      message: "Bank PDF import is already running."
    };
  }

  try {
    const startedAt = new Date();
    const allClients = getEnabledClients_();
    const clientIdFilter = String(options.clientId || "").trim();

    const clients = clientIdFilter
      ? allClients.filter(function(c) {
          return c.clientId === clientIdFilter;
        })
      : allClients;

    if (clientIdFilter && clients.length === 0) {
      throw new Error("No enabled client found with clientId=" + clientIdFilter);
    }

    const result = {
      success: true,
      mode: mode,
      startedAt: startedAt.toISOString(),
      finishedAt: "",
      clientCount: clients.length,
      totals: {
        checkedThreads: 0,
        checkedMessages: 0,
        checkedAttachments: 0,
        savedFiles: 0,
        skippedExisting: 0,
        skippedNonPdf: 0,
        skippedSender: 0,
        errors: 0
      },
      clients: []
    };

    for (let i = 0; i < clients.length; i++) {
      const client = clients[i];

      try {
        const clientResult = processOneClient_(client, mode, options);
        result.clients.push(clientResult);

        addTotals_(result.totals, clientResult);

      } catch (clientErr) {
        result.totals.errors++;

        result.clients.push({
          success: false,
          clientId: client.clientId,
          code: "CLIENT_ERROR",
          message: clientErr && clientErr.message
            ? clientErr.message
            : String(clientErr)
        });
      }
    }

    result.finishedAt = new Date().toISOString();

    console.log(JSON.stringify(result, null, 2));
    return result;

  } finally {
    lock.releaseLock();
  }
}


function processOneClient_(client, mode, options) {
  const folder = DriveApp.getFolderById(client.driveFolderId);
  const query = buildGmailQuery_(client, mode, options);

  const result = {
    success: true,
    clientId: client.clientId,
    mode: mode,
    folderName: folder.getName(),
    query: query,
    checkedThreads: 0,
    checkedMessages: 0,
    checkedAttachments: 0,
    savedFiles: 0,
    skippedExisting: 0,
    skippedNonPdf: 0,
    skippedSender: 0,
    errors: 0
  };

  let start = 0;
  const maxThreads = client.maxThreadsPerRun;
  const pageSize = client.pageSize;

  while (start < maxThreads) {
    const remaining = maxThreads - start;
    const fetchSize = Math.min(pageSize, remaining);

    const threads = GmailApp.search(query, start, fetchSize);

    if (!threads || threads.length === 0) {
      break;
    }

    for (let i = 0; i < threads.length; i++) {
      processOneThread_(threads[i], client, folder, result);
    }

    if (threads.length < fetchSize) {
      break;
    }

    start += fetchSize;
  }

  return result;
}


function processOneThread_(thread, client, folder, result) {
  result.checkedThreads++;

  const messages = thread.getMessages();
  let savedSomethingFromThread = false;

  for (let i = 0; i < messages.length; i++) {
    const msg = messages[i];
    result.checkedMessages++;

    if (!isAllowedSender_(msg, client)) {
      result.skippedSender++;
      continue;
    }

    const attachments = msg.getAttachments({
      includeInlineImages: false,
      includeAttachments: true
    });

    for (let k = 0; k < attachments.length; k++) {
      const attachment = attachments[k];
      const originalName = String(attachment.getName() || "").trim();

      if (!isPdfAttachment_(attachment, originalName)) {
        result.skippedNonPdf++;
        continue;
      }

      result.checkedAttachments++;

      const finalName = buildStableFileName_(client, msg, originalName, k + 1);

      if (driveFileExistsByName_(folder, finalName)) {
        result.skippedExisting++;
        continue;
      }

      const blob = attachment.copyBlob().setName(finalName);
      folder.createFile(blob);

      result.savedFiles++;
      savedSomethingFromThread = true;
    }
  }

  if (savedSomethingFromThread && client.markReadAfterSave) {
    thread.markRead();
  }

  if (savedSomethingFromThread && client.archiveAfterSave) {
    thread.moveToArchive();
  }
}


/* ======================================================
 * Client config
 * ====================================================== */

function getEnabledClients_() {
  const props = PropertiesService.getScriptProperties();

  const raw = String(props.getProperty(BANK_PROP_CLIENTS_JSON) || "").trim();

  if (!raw) {
    throw new Error(
      "Missing Script Property " +
      BANK_PROP_CLIENTS_JSON +
      ". Add JSON array of client configs."
    );
  }

  let parsed;

  try {
    parsed = JSON.parse(raw);
  } catch (err) {
    throw new Error(
      BANK_PROP_CLIENTS_JSON +
      " is not valid JSON. " +
      (err && err.message ? err.message : String(err))
    );
  }

  if (!Array.isArray(parsed)) {
    throw new Error(BANK_PROP_CLIENTS_JSON + " must be a JSON array.");
  }

  const clients = [];

  for (let i = 0; i < parsed.length; i++) {
    const client = normalizeClientConfig_(parsed[i], i + 1);

    if (client.enabled) {
      clients.push(client);
    }
  }

  if (clients.length === 0) {
    throw new Error("No enabled clients in " + BANK_PROP_CLIENTS_JSON + ".");
  }

  return clients;
}


function normalizeClientConfig_(rawClient, index) {
  const c = rawClient || {};

  const client = {
    clientId: String(c.clientId || ("CLIENT_" + index)).trim(),
    enabled: c.enabled !== false,

    driveFolderId: String(c.driveFolderId || "").trim(),

    bankSenders: Array.isArray(c.bankSenders)
      ? c.bankSenders.map(function(x) {
          return String(x || "").trim();
        }).filter(Boolean)
      : [],

    searchDays: toPositiveInt_(c.searchDays, 30),
    maxThreadsPerRun: toPositiveInt_(c.maxThreadsPerRun, 100),
    pageSize: toPositiveInt_(c.pageSize, 50),

    gmailQueryExtra: String(c.gmailQueryExtra || "").trim(),

    fileNamePrefix: String(c.fileNamePrefix || "").trim(),

    archiveAfterSave: c.archiveAfterSave === true,
    markReadAfterSave: c.markReadAfterSave === true
  };

  validateClientConfig_(client, index);

  return client;
}


function validateClientConfig_(client, index) {
  if (!client.clientId) {
    throw new Error("Client #" + index + " has empty clientId.");
  }

  if (!client.driveFolderId) {
    throw new Error(
      "Client " + client.clientId + " has empty driveFolderId."
    );
  }

  if (client.driveFolderId.indexOf("https://") === 0) {
    throw new Error(
      "Client " +
      client.clientId +
      " driveFolderId contains full URL. Use only folder ID."
    );
  }

  if (client.bankSenders.length === 0 && !client.gmailQueryExtra) {
    throw new Error(
      "Client " +
      client.clientId +
      " must have bankSenders or gmailQueryExtra."
    );
  }

  if (client.pageSize > 500) {
    throw new Error(
      "Client " + client.clientId + " pageSize too large. Use <= 500."
    );
  }

  if (client.maxThreadsPerRun > 1000) {
    throw new Error(
      "Client " +
      client.clientId +
      " maxThreadsPerRun too large. Use <= 1000."
    );
  }
}


/* ======================================================
 * Gmail query
 * ====================================================== */

function buildGmailQuery_(client, mode, options) {
  const parts = [
    "in:inbox",
    "has:attachment",
    "filename:pdf"
  ];

  if (client.bankSenders.length > 0) {
    const senderQuery = client.bankSenders
      .map(function(email) {
        return "from:" + email;
      })
      .join(" OR ");

    parts.push("(" + senderQuery + ")");
  }

  if (mode === "backfill" || mode === "webapp-backfill") {
    const fromDate = String(options.fromDate || "").trim();
    const toDate = String(options.toDate || "").trim() || getTodayIsoDate_();

    if (!fromDate) {
      throw new Error("Backfill mode requires fromDate.");
    }

    assertIsoDate_(fromDate, "fromDate");
    assertIsoDate_(toDate, "toDate");

    const beforeExclusive = addDaysIso_(toDate, 1);

    parts.push("after:" + gmailDate_(fromDate));
    parts.push("before:" + gmailDate_(beforeExclusive));

  } else {
    parts.push("newer_than:" + client.searchDays + "d");
  }

  if (client.gmailQueryExtra) {
    parts.push("(" + client.gmailQueryExtra + ")");
  }

  return parts.join(" ");
}


function isAllowedSender_(msg, client) {
  if (!client.bankSenders || client.bankSenders.length === 0) {
    return true;
  }

  const from = String(msg.getFrom() || "").toLowerCase();

  return client.bankSenders.some(function(email) {
    return from.indexOf(String(email).toLowerCase()) !== -1;
  });
}


/* ======================================================
 * File handling
 * ====================================================== */

function isPdfAttachment_(attachment, fileName) {
  const name = String(fileName || "").toLowerCase();

  if (name.endsWith(".pdf")) {
    return true;
  }

  const contentType = String(attachment.getContentType() || "").toLowerCase();

  return contentType === "application/pdf";
}


function buildStableFileName_(client, msg, originalName, attachmentIndex) {
  const datePart = Utilities.formatDate(
    msg.getDate(),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd"
  );

  const messageId = sanitizeFileNamePart_(msg.getId()).slice(0, 32);

  let safeOriginal = sanitizeFileNamePart_(originalName);

  if (!safeOriginal.toLowerCase().endsWith(".pdf")) {
    safeOriginal += ".pdf";
  }

  const parts = [];

  if (client.fileNamePrefix) {
    parts.push(sanitizeFileNamePart_(client.fileNamePrefix));
  }

  parts.push(datePart);
  parts.push(messageId);
  parts.push("att" + String(attachmentIndex));
  parts.push(safeOriginal);

  return parts.join("_");
}


function driveFileExistsByName_(folder, fileName) {
  const files = folder.getFilesByName(fileName);
  return files.hasNext();
}


function sanitizeFileNamePart_(value) {
  return String(value || "file")
    .replace(/[\\\/:*?"<>|#%{}~&]/g, "_")
    .replace(/\s+/g, "_")
    .replace(/_+/g, "_")
    .trim();
}


/* ======================================================
 * WebApp helpers
 * ====================================================== */

function parsePostJson_(e) {
  const raw = e && e.postData && e.postData.contents
    ? e.postData.contents
    : "{}";

  try {
    return JSON.parse(raw);
  } catch (err) {
    throw new Error("Invalid JSON body.");
  }
}


function validateWebAppSecret_(data) {
  const props = PropertiesService.getScriptProperties();

  const expected = String(props.getProperty(BANK_PROP_SECRET) || "").trim();

  if (!expected) {
    throw new Error(
      "Missing Script Property " +
      BANK_PROP_SECRET +
      ". WebApp calls are disabled until secret is configured."
    );
  }

  const actual = String(data.secret || "").trim();

  if (actual !== expected) {
    throw new Error("Invalid BANK_IMPORT_SECRET.");
  }
}


function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


/* ======================================================
 * Utility helpers
 * ====================================================== */

function toPositiveInt_(value, fallback) {
  const n = parseInt(value, 10);

  if (!isFinite(n) || n <= 0) {
    return fallback;
  }

  return n;
}


function addTotals_(totals, clientResult) {
  totals.checkedThreads += Number(clientResult.checkedThreads || 0);
  totals.checkedMessages += Number(clientResult.checkedMessages || 0);
  totals.checkedAttachments += Number(clientResult.checkedAttachments || 0);
  totals.savedFiles += Number(clientResult.savedFiles || 0);
  totals.skippedExisting += Number(clientResult.skippedExisting || 0);
  totals.skippedNonPdf += Number(clientResult.skippedNonPdf || 0);
  totals.skippedSender += Number(clientResult.skippedSender || 0);
  totals.errors += Number(clientResult.errors || 0);
}


function assertIsoDate_(value, fieldName) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(value || ""))) {
    throw new Error(fieldName + " must be in YYYY-MM-DD format.");
  }
}


function gmailDate_(isoDate) {
  assertIsoDate_(isoDate, "gmailDate");
  return isoDate.replace(/-/g, "/");
}


function getTodayIsoDate_() {
  return Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd"
  );
}


function addDaysIso_(isoDate, days) {
  assertIsoDate_(isoDate, "isoDate");

  const parts = isoDate.split("-");
  const d = new Date(
    Number(parts[0]),
    Number(parts[1]) - 1,
    Number(parts[2])
  );

  d.setDate(d.getDate() + days);

  return Utilities.formatDate(
    d,
    Session.getScriptTimeZone(),
    "yyyy-MM-dd"
  );
}


/* ======================================================
 * Optional helper: print example properties
 * ====================================================== */

function printExampleBankImportProperties() {
  const example = [
    {
      clientId: "CLIENT_A",
      enabled: true,
      driveFolderId: "PASTE_SHARED_DRIVE_FOLDER_ID_HERE",
      bankSenders: [
        "izvodi@banka.rs",
        "noreply@banka.rs"
      ],
      searchDays: 30,
      maxThreadsPerRun: 100,
      pageSize: 50,
      gmailQueryExtra: "",
      fileNamePrefix: "CLIENT_A",
      archiveAfterSave: false,
      markReadAfterSave: false
    }
  ];

  console.log(BANK_PROP_CLIENTS_JSON + " =");
  console.log(JSON.stringify(example, null, 2));

  console.log(BANK_PROP_BACKFILL_FROM_DATE + " = 2026-01-01");
  console.log(BANK_PROP_BACKFILL_TO_DATE + " = " + getTodayIsoDate_());
  console.log(BANK_PROP_SECRET + " = neki-dug-random-secret");
}
