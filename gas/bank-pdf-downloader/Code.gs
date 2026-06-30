/*******************************************************
 * AgriX_C00X_BankPdfDownloader
 *
 * Standalone Apps Script project, installed on the Google
 * account that RECEIVES bank statement emails.
 *
 * Single responsibility:
 *   Gmail bank PDF attachments  ->  shared Drive folder (Bank_Izvodi)
 *
 * It does NOT parse PDF, does NOT write to tblBankaImport,
 * does NOT touch tblNovac. That whole downstream chain already
 * lives in the Excel/VBA client and stays unchanged:
 *   Drive folder (Bank_Izvodi)
 *     -> PullBankPdfsFromDriveProduction  (modBankaImport.bas)
 *     -> ImportBankaInbox_TX               -> tblBankaImport
 *     -> modBankaMapiranje                 -> tblNovac
 *
 * Deployment model (project decision):
 *   - separate GAS project, NOT folded into AgriX_C00X_GAS_PROD
 *     (so there is no doPost collision with gas/Code.gs)
 *   - daily time-trigger only; no Web App / doPost / secret
 *
 * Per-client config lives in Script Properties so the SAME source
 * deploys to every client (C001, C002, ...) without source edits,
 * matching the folder-id-via-properties pattern in gas/DriveFolder.gs:
 *
 *   AGRIX_BANK_IZVODI_FOLDER_ID  (required)
 *       Drive folder ID of the Bank_Izvodi folder that syncs down
 *       (Google Drive for Desktop) to the VBA bank inbox path
 *       (BANKA_INBOX_PATH / BANKA_DRIVE_SOURCE_PATH). Use the SAME
 *       folder ID the main project stores under this property name.
 *       Share that folder with this Gmail account as Editor.
 *
 *   BANK_PDF_SENDERS             (required)
 *       Bank sender email addresses, comma / semicolon / whitespace
 *       separated, e.g. "izvodi@banka.rs, noreply@komercijalna.rs".
 *******************************************************/

// Non-secret tuning only. Folder ID and senders come from Script
// Properties (resolveBankPdfRuntimeConfig_), never hardcoded here.
const BANK_PDF_CONFIG = {
  // Script Property names that hold the per-client values.
  folderIdProp: 'AGRIX_BANK_IZVODI_FOLDER_ID',
  sendersProp: 'BANK_PDF_SENDERS',

  // Gmail label applied to a thread after at least one PDF was saved.
  // The search query excludes it, so each statement is downloaded
  // exactly once even though the VBA side later MOVES the saved file
  // out of the folder (Downloaded/Verarbeitet). This is the robust
  // dedupe; the per-file name check below is only a secondary guard.
  processedLabel: 'AgriX-BankPdfSaved',

  // Search only recent mail. For a daily run, 30 is safe.
  searchDays: 30,

  // Batch size. Keep small and safe.
  maxThreadsPerRun: 50,

  // Optional: after a thread saved something, also mark it read.
  markReadAfterSave: false,

  // Optional: after a thread saved something, also archive it.
  archiveAfterSave: false
};


/**
 * Main daily/manual function. Run by the daily trigger
 * (setupDailyBankPdfDownloader) or manually from the editor.
 */
function saveBankPdfsToProjectDrive() {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(5000)) {
    return {
      success: false,
      code: 'LOCKED',
      message: 'Bank PDF downloader is already running.'
    };
  }

  try {
    const cfg = resolveBankPdfRuntimeConfig_();

    const folder = DriveApp.getFolderById(cfg.folderId);
    const label = getOrCreateLabel_(BANK_PDF_CONFIG.processedLabel);
    const query = buildBankPdfGmailQuery_(cfg.senders);

    const threads = GmailApp.search(query, 0, BANK_PDF_CONFIG.maxThreadsPerRun);

    let checkedThreads = 0;
    let checkedMessages = 0;
    let checkedAttachments = 0;
    let savedFiles = 0;
    let skippedExisting = 0;
    let skippedNonPdf = 0;
    let labeledThreads = 0;

    for (let i = 0; i < threads.length; i++) {
      const thread = threads[i];
      checkedThreads++;

      const messages = thread.getMessages();
      let threadSavedSomething = false;

      for (let j = 0; j < messages.length; j++) {
        const msg = messages[j];
        checkedMessages++;

        if (!isAllowedBankSender_(msg, cfg.senders)) {
          continue;
        }

        const attachments = msg.getAttachments({
          includeInlineImages: false,
          includeAttachments: true
        });

        for (let k = 0; k < attachments.length; k++) {
          const attachment = attachments[k];
          const originalName = String(attachment.getName() || '').trim();

          if (!isPdfAttachment_(attachment, originalName)) {
            skippedNonPdf++;
            continue;
          }

          checkedAttachments++;

          const finalName = buildStableDriveFileName_(msg, originalName, k + 1);

          if (driveFileExistsByName_(folder, finalName)) {
            skippedExisting++;
            continue;
          }

          const blob = attachment.copyBlob().setName(finalName);
          folder.createFile(blob);

          savedFiles++;
          threadSavedSomething = true;
        }
      }

      if (threadSavedSomething) {
        thread.addLabel(label);
        labeledThreads++;

        if (BANK_PDF_CONFIG.markReadAfterSave) {
          thread.markRead();
        }
        if (BANK_PDF_CONFIG.archiveAfterSave) {
          thread.moveToArchive();
        }
      }
    }

    const result = {
      success: true,
      query: query,
      checkedThreads: checkedThreads,
      checkedMessages: checkedMessages,
      checkedAttachments: checkedAttachments,
      savedFiles: savedFiles,
      skippedExisting: skippedExisting,
      skippedNonPdf: skippedNonPdf,
      labeledThreads: labeledThreads,
      runAt: new Date().toISOString()
    };

    console.log(JSON.stringify(result, null, 2));
    return result;

  } catch (err) {
    const result = {
      success: false,
      code: 'ERROR',
      message: err && err.message ? err.message : String(err),
      runAt: new Date().toISOString()
    };

    console.error(JSON.stringify(result, null, 2));
    return result;

  } finally {
    lock.releaseLock();
  }
}


/**
 * Daily trigger setup. Run this once manually from the editor.
 * Idempotent: removes any existing trigger for the handler first.
 */
function setupDailyBankPdfDownloader() {
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    if (trigger.getHandlerFunction() === 'saveBankPdfsToProjectDrive') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('saveBankPdfsToProjectDrive')
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .create();
}


/**
 * Quick config test from the editor: resolves Script Properties,
 * opens the folder, builds the query, counts a few sample threads.
 */
function testBankPdfDownloaderConfig() {
  const cfg = resolveBankPdfRuntimeConfig_();

  const folder = DriveApp.getFolderById(cfg.folderId);
  const query = buildBankPdfGmailQuery_(cfg.senders);
  const threads = GmailApp.search(query, 0, 5);

  const result = {
    success: true,
    folderName: folder.getName(),
    senders: cfg.senders,
    query: query,
    sampleThreadCount: threads.length
  };

  console.log(JSON.stringify(result, null, 2));
  return result;
}


/**
 * Minimal Gmail-scope smoke test (no Drive, no properties).
 */
function testGmailAccessOnly() {
  const threads = GmailApp.search('in:inbox newer_than:30d', 0, 1);

  console.log('Gmail OK. Threads found: ' + threads.length);

  return {
    success: true,
    threadCount: threads.length
  };
}


/* =========================
 * Internal helpers
 * ========================= */

/**
 * Reads and validates per-client config from Script Properties.
 * Throws with an actionable message if anything required is missing.
 */
function resolveBankPdfRuntimeConfig_() {
  const props = PropertiesService.getScriptProperties();

  const folderId = String(props.getProperty(BANK_PDF_CONFIG.folderIdProp) || '').trim();
  if (!folderId) {
    throw new Error(
      'Missing Script Property ' + BANK_PDF_CONFIG.folderIdProp +
      ' (Drive folder ID of Bank_Izvodi that syncs to the VBA bank inbox).'
    );
  }

  const senders = parseSenders_(props.getProperty(BANK_PDF_CONFIG.sendersProp));
  if (!senders.length) {
    throw new Error(
      'Missing Script Property ' + BANK_PDF_CONFIG.sendersProp +
      ' (bank sender emails, comma separated).'
    );
  }

  if (!BANK_PDF_CONFIG.searchDays || BANK_PDF_CONFIG.searchDays < 1) {
    throw new Error('searchDays must be >= 1.');
  }

  if (!BANK_PDF_CONFIG.maxThreadsPerRun || BANK_PDF_CONFIG.maxThreadsPerRun < 1) {
    throw new Error('maxThreadsPerRun must be >= 1.');
  }

  return { folderId: folderId, senders: senders };
}

function parseSenders_(raw) {
  return String(raw || '')
    .split(/[,;\s]+/)
    .map(function (s) { return s.trim().toLowerCase(); })
    .filter(function (s) { return s.length > 0; });
}

function getOrCreateLabel_(name) {
  return GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name);
}

function buildBankPdfGmailQuery_(senders) {
  const fromClause = senders
    .map(function (email) { return 'from:' + email; })
    .join(' OR ');

  return [
    'in:inbox',
    'has:attachment',
    'filename:pdf',
    'newer_than:' + BANK_PDF_CONFIG.searchDays + 'd',
    '-label:' + BANK_PDF_CONFIG.processedLabel,
    '(' + fromClause + ')'
  ].join(' ');
}

function isAllowedBankSender_(msg, senders) {
  const from = String(msg.getFrom() || '').toLowerCase();

  return senders.some(function (email) {
    return from.indexOf(email) !== -1;
  });
}

function isPdfAttachment_(attachment, fileName) {
  const name = String(fileName || '').toLowerCase();

  if (name.endsWith('.pdf')) {
    return true;
  }

  const contentType = String(attachment.getContentType() || '').toLowerCase();
  return contentType === 'application/pdf';
}

/**
 * Stable, collision-free Drive file name. Always ends in ".pdf" so the
 * VBA puller (Dir$ "*.pdf") will see it even when the original Gmail
 * attachment name has no extension.
 */
function buildStableDriveFileName_(msg, originalName, attachmentIndex) {
  const datePart = Utilities.formatDate(
    msg.getDate(),
    Session.getScriptTimeZone(),
    'yyyy-MM-dd'
  );

  const messageId = sanitizeFileNamePart_(msg.getId()).slice(0, 32);
  let safeOriginalName = sanitizeFileNamePart_(originalName);

  if (!/\.pdf$/i.test(safeOriginalName)) {
    safeOriginalName = safeOriginalName + '.pdf';
  }

  return [
    datePart,
    messageId,
    'att' + String(attachmentIndex),
    safeOriginalName
  ].join('_');
}

function driveFileExistsByName_(folder, fileName) {
  return folder.getFilesByName(fileName).hasNext();
}

function sanitizeFileNamePart_(value) {
  return String(value || 'file')
    .replace(/[\\\/:*?"<>|#%{}~&]/g, '_')
    .replace(/\s+/g, '_')
    .replace(/_+/g, '_')
    .trim();
}
