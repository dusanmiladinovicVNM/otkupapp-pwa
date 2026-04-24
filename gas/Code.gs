// ============================================================
// OtkupApp - Google Apps Script Backend
// Deployed as Web App: doPost receives JSON from PWA
// Writes to Google Sheet "OTK-{otkupacName}"
//
// Setup:
// 1. Erstelle neues Google Apps Script Projekt (script.google.com)
// 2. Füge diesen Code als Code.gs ein
// 3. Setze MASTER_FOLDER_ID auf die Google Drive Folder ID
// 4. Deploy → Web App → Execute as: Me, Access: Anyone
// 5. Kopiere die Web App URL in die PWA Config
// ============================================================

// --- CONFIG ---
const MASTER_FOLDER_ID = '1EUNbItjRoqSPaW-ZDSn-JD1x1CuYSe26'; // Google Drive Ordner für alle OTK-Sheets

const COLUMNS = [
  'ClientRecordID',
  'ServerRecordID',
  'CreatedAtClient',
  'UpdatedAtClient',
  'UpdatedAtServer',
  'SyncStatus',
  'DeviceID',
  'OtkupacID',
  'Datum',
  'KooperantID',
  'KooperantName',
  'VrstaVoca',
  'SortaVoca',
  'Klasa',
  'Kolicina',
  'Cena',
  'TipAmbalaze',
  'KolAmbalaze',
  'ParcelaID',
  'VozacID',
  'Napomena',
  'ReceivedAt'
];

const ZBIRNA_COLUMNS = [
  'ClientRecordID',
  'ServerRecordID',
  'CreatedAtClient',
  'UpdatedAtClient',
  'UpdatedAtServer',
  'SyncStatus',
  'VozacID',
  'Datum',
  'KupacID',
  'KupacName',
  'VrstaVoca',
  'SortaVoca',
  'KolicinaKlI',
  'KolicinaKlII',
  'TipAmbalaze',
  'KolAmbalaze',
  'Klasa',
  'OtkupRecordIDs',
  'ReceivedAt',
  'BrojZbirne'
];

const AGROMERE_COLUMNS = [
  'ClientRecordID','CreatedAtClient','SyncStatus','KooperantID','ParcelaID',
  'Mera','Datum','Vreme','Napomena','ReceivedAt'
];

const WARROOM_DEMAND_COLUMNS = [
  'DemandID', 'Datum', 'KupacID', 'KupacName', 'Kg', 'Vrsta', 'Klasa', 'Primljeno', 'CreatedAt'
];

const KAMION_STATUS_COLUMNS = [
  'VozacID', 'Status', 'Ruta', 'UpdatedAt'
];

const DISPECER_PLAN_COLUMNS = [
    'PlanID','Datum','DemandID',
    'VozacID','VozacName',
    'StanicaID','StanicaName',
    'KupacID','KupacName',
    'PlannedKg','Status',
    'CreatedAt','UpdatedAt'
];

const IZDAVANJE_COLUMNS = [
  'IzdavanjeID','Datum','KooperantID','KooperantName','ParcelaID',
  'Stavke','UkupnaVrednost','IzdaoUser','Napomena',
  'SigIzdavalac','SigPrimalac','CreatedAt'
];

const TRETMAN_COLUMNS = [
  'ClientRecordID',
  'ServerRecordID',
  'CreatedAtClient',
  'UpdatedAtClient',
  'UpdatedAtServer',
  'SyncStatus',
  'KooperantID',
  'ParcelaID',
  'Datum',
  'Mera',
  'ArtikalID',
  'ArtikalNaziv',
  'KolicinaUpotrebljena',
  'JedinicaMere',
  'DozaPreporucena',
  'DozaPrimenjena',
  'OpremaTraktor',
  'OpremaPrskalica',
  'OpremaOstalo',
  'KarencaDana',
  'DatumBerbeDozvoljeno',
  'VremePocetka',
  'VremeZavrsetka',
  'TrajanjeMinuta',
  'GeoLatStart',
  'GeoLngStart',
  'GeoLatEnd',
  'GeoLngEnd',
  'GeoAutoDetect',
  'MeteoTemp',
  'MeteoWind',
  'MeteoHumidity',
  'MeteoOverride',
  'Napomena',
  'ReceivedAt'
];

const OPREMA_COLUMNS = [
  'ClientRecordID',
  'ServerRecordID',
  'CreatedAtClient',
  'UpdatedAtClient',
  'UpdatedAtServer',
  'SyncStatus',
  'KooperantID',
  'Naziv',
  'Tip',
  'ReceivedAt'
];

const TROSKOVI_COLUMNS = [
  'ClientRecordID','CreatedAtClient','SyncStatus','KooperantID','ParcelaID',
  'Datum','Kategorija','Opis','Iznos','DokumentBroj','Napomena','ReceivedAt'
];

const FISKALNI_COLUMNS = [
  'ClientRecordID','CreatedAtClient','SyncStatus','KooperantID',
  'InvoiceNumber','Kompanija','DatumRacuna','VerificationUrl',
  'NazivStavke','ArtikalID','ArtikalNaziv','Kolicina','JedCena','Ukupno',
  'PDVStopa','Mapirano','ReceivedAt'
];

const FISKALNI_MAP_COLUMNS = [
  'FiskalniNaziv','ArtikalID','ArtikalNaziv','KooperantID','CreatedAt'
];

const GEO_SPREADSHEET_ID = '1hOkvtcHhnGXc5FKv9gG6lRMvxxGK1rOEYo8EjZPJG24';
const GEO_SHEET_PARCELE = 'Parcele';
// ============================================================
// ENDPOINTS
// ============================================================

function getJsonBody(e) {
    try {
        if (!e || !e.postData || !e.postData.contents) return {};
        return JSON.parse(e.postData.contents || '{}') || {};
    } catch (err) {
        return {};
    }
}

function handleAuthorizedRead(data, tokenData) {
  const action = data.action || '';

  if (action === 'getStammdaten') {
    return jsonResponse(getStammdaten());
  }

  if (action === 'getOtkupi') {
    const otkupacID = data.otkupacID || '';
    if (tokenData.entityID !== otkupacID) {
      return jsonResponse({ success: false, error: 'Nemate pristup', code: 403 });
    }
    return jsonResponse(getOtkupiForOtkupac(otkupacID));
  }

  if (action === 'getKartica') {
    const kooperantID = data.kooperantID || '';
    if (tokenData.entityID !== kooperantID) {
      return jsonResponse({ success: false, error: 'Nemate pristup', code: 403 });
    }
    return jsonResponse(getKarticaForKooperant(kooperantID));
  }

  if (action === 'getAgromere') {
    const kooperantID = data.kooperantID || '';
    if (tokenData.entityID !== kooperantID) {
      return jsonResponse({ success: false, error: 'Nemate pristup', code: 403 });
    }
    return jsonResponse(getAgromereForKooperant(kooperantID));
  }

  if (action === 'getMgmtKartica') {
    if (tokenData.role !== 'Management') {
      return jsonResponse({ success: false, error: 'Nemate pristup', code: 403 });
    }
    const kooperantID = data.kooperantID || '';
    return jsonResponse(getKarticaForKooperant(kooperantID));
  }

  if (action === 'getMgmtOtkupiByStanica') {
    if (tokenData.role !== 'Management') {
      return jsonResponse({ success: false, error: 'Nemate pristup', code: 403 });
    }
    const stanicaID = data.stanicaID || '';
    return jsonResponse(getOtkupiByStanica(stanicaID));
  }

  if (action === 'getMgmtSaldoOM') {
    if (tokenData.role !== 'Management') {
      return jsonResponse({ success: false, error: 'Nemate pristup', code: 403 });
    }
    return jsonResponse(getSaldoOM());
  }

  if (action === 'getMgmtSaldoKupci') {
    if (tokenData.role !== 'Management') {
      return jsonResponse({ success: false, error: 'Nemate pristup', code: 403 });
    }
    return jsonResponse(getSaldoKupci());
  }

  if (action === 'getMgmtOtkupPoOM') {
    if (tokenData.role !== 'Management') {
      return jsonResponse({ success: false, error: 'Nemate pristup', code: 403 });
    }
    return jsonResponse(getMgmtReport('OtkupPoOM'));
  }

  if (action === 'getMgmtPredatoPoKupcu') {
    if (tokenData.role !== 'Management') {
      return jsonResponse({ success: false, error: 'Nemate pristup', code: 403 });
    }
    return jsonResponse(getMgmtReport('PredatoPoKupcu'));
  }

  if (action === 'getMgmtAll') {
    if (tokenData.role !== 'Management') {
      return jsonResponse({ success: false, error: 'Nemate pristup', code: 403 });
    }
    return jsonResponse({
      success: true,
      saldoOM: getMgmtReport('SaldoOM').records || [],
      saldoKupci: getMgmtReport('SaldoKupci').records || [],
      otkupPoOM: getMgmtReport('OtkupPoOM').records || [],
      predatoPoKupcu: getMgmtReport('PredatoPoKupcu').records || [],
      kartice: getKarticaAll().records || [],
      saldoOMDetail: getMgmtReport('SaldoOMDetail').records || [],
      fakture: getMgmtReport('Fakture').records || [],
      fakturaStavke: getMgmtReport('FakturaStavke').records || [],
      otkupiAll: getAllOtkupiSheets()
    });
  }

  if (action === 'getMgmtFakture') {
    if (tokenData.role !== 'Management') {
      return jsonResponse({ success: false, error: 'Nemate pristup', code: 403 });
    }
    const kupacID = data.kupacID || '';
    return jsonResponse(getFaktureByKupac(kupacID));
  }

  if (action === 'getMgmtFakturaStavke') {
    if (tokenData.role !== 'Management') {
      return jsonResponse({ success: false, error: 'Nemate pristup', code: 403 });
    }
    const fakturaID = data.fakturaID || '';
    return jsonResponse(getFakturaStavke(fakturaID));
  }

  if (action === 'getVozacOtkupi') {
    if (tokenData.role !== 'Vozac') {
      return jsonResponse({ success: false, error: 'Nemate pristup', code: 403 });
    }
    const vozacID = tokenData.entityID;
    return jsonResponse(getOtkupiForVozac(vozacID));
  }

  if (action === 'getVozacZbirne') {
    if (tokenData.role !== 'Vozac') {
      return jsonResponse({ success: false, error: 'Nemate pristup', code: 403 });
    }
    const vozacID = tokenData.entityID;
    return jsonResponse(getZbirneForVozac(vozacID));
  }

  if (action === 'getWarRoomDemand') {
    if (tokenData.role !== 'Management') {
      return jsonResponse({ success: false, error: 'Nemate pristup', code: 403 });
    }
    return jsonResponse(getWarRoomDemand());
  }

  if (action === 'getDispecer') {
    if (tokenData.role !== 'Management') {
      return jsonResponse({ success: false, error: 'Nemate pristup', code: 403 });
    }
    return jsonResponse(getDispecer());
  }

  if (action === 'getKamionStatus') {
    if (tokenData.role !== 'Management') {
      return jsonResponse({ success: false, error: 'Nemate pristup', code: 403 });
    }
    return jsonResponse(getKamionStatus());
  }

  if (action === 'getTretmani') {
    const kooperantID = data.kooperantID || '';
    if (tokenData.entityID !== kooperantID) {
      return jsonResponse({ success: false, error: 'Nemate pristup', code: 403 });
    }
    return jsonResponse(getTretmaniForKooperant(kooperantID));
  }

  if (action === 'getOprema') {
    const kooperantID = data.kooperantID || '';
    if (tokenData.entityID !== kooperantID) {
      return jsonResponse({ success: false, error: 'Nemate pristup', code: 403 });
    }
    return jsonResponse(getOpremaForKooperant(kooperantID));
  }

  if (action === 'getKooperantProizvodnja') {
    const kooperantID = data.kooperantID || '';
    if (tokenData.entityID !== kooperantID) {
      return jsonResponse({ success: false, error: 'Nemate pristup', code: 403 });
    }
    return jsonResponse(getKooperantProizvodnja(kooperantID));
  }

  if (action === 'getTroskovi') {
    const kooperantID = data.kooperantID || '';
    if (tokenData.entityID !== kooperantID) {
      return jsonResponse({ success: false, error: 'Nemate pristup', code: 403 });
    }
    return jsonResponse(getTroskoviForKooperant(kooperantID));
  }

  return null;
}

function handlePublicRead(data) {
  const action = data.action || '';

  if (action === 'getParcelGeo') {
    return jsonResponse(getParcelGeo(data.parcelaId));
  }

  if (action === 'getParcelMeteo') {
    return jsonResponse(getParcelMeteo(data.parcelaId));
  }

  if (action === 'getParcelMeteoLatest') {
    return jsonResponse(getParcelMeteoLatest(data.parcelaId));
  }

  if (action === 'getAllMeteoLatest') {
    return jsonResponse(getAllMeteoLatest());
  }

  return null;
}

function doPost(e) {
  try {
    const data = getJsonBody(e);

    const publicReadResponse = handlePublicRead(data);
    if (publicReadResponse) return publicReadResponse;

    if (data.action === 'login') {
      return jsonResponse(authenticateUser(data.username, data.pin));
    }

    // ostaje javno samo ako to namerno želiš;
    // u suprotnom prebaci ispod auth check-a
    if (data.action === 'saveParcelPolygon') {
      return jsonResponse(saveParcelPolygon(data));
    }

    if (!validateToken(data.token)) {
      return jsonResponse({ success: false, error: 'Neautorizovan pristup', code: 401 });
    }

    const tokenData = getTokenData(data.token);
    const readResponse = handleAuthorizedRead(data, tokenData);
    if (readResponse) return readResponse;

    if (data.action === 'sync') {
      if (!Array.isArray(data.records)) {
        return jsonResponse({ success: false, error: 'records must be an array' });
      }
      const lockResult = withLock(function() {
        const results = data.records.map(r => processRecord(r, data.otkupacID));
        return {
          success: true,
          processed: results.length,
          succeeded: results.filter(r => r && r.success).length,
          failed: results.filter(r => !r || !r.success).length,
          results: results
        };
      });
      return jsonResponse(lockResult);
    }

    if (data.action === 'syncAgromere') {
      if (!Array.isArray(data.records)) {
        return jsonResponse({ success: false, error: 'records must be an array' });
      }
      const lockResult = withLock(function() {
        const results = data.records.map(r => processAgromereRecord(r, data.kooperantID));
        return {
          success: true,
          processed: results.length,
          succeeded: results.filter(r => r && r.success).length,
          failed: results.filter(r => !r || !r.success).length,
          results: results
        };
      });
      return jsonResponse(lockResult);
    }

    if (data.action === 'syncZbirna') {
      if (!Array.isArray(data.records)) {
        return jsonResponse({ success: false, error: 'records must be an array' });
      }
      const lockResult = withLock(function() {
        const results = data.records.map(r => processZbirnaRecord(r, data.vozacID));
        return {
          success: true,
          processed: results.length,
          succeeded: results.filter(r => r && r.success).length,
          failed: results.filter(r => !r || !r.success).length,
          results: results
        };
      });
      return jsonResponse(lockResult);
    }

    if (data.action === 'syncTretman') {
      if (!Array.isArray(data.records)) {
        return jsonResponse({ success: false, error: 'records must be an array' });
      }
      const lockResult = withLock(function() {
        const results = data.records.map(r => processTretmanRecord(r, data.kooperantID));
        return {
          success: true,
          processed: results.length,
          succeeded: results.filter(r => r && r.success).length,
          failed: results.filter(r => !r || !r.success).length,
          results: results
        };
      });
      return jsonResponse(lockResult);
    }

    if (data.action === 'syncOprema') {
      if (!Array.isArray(data.records)) {
        return jsonResponse({ success: false, error: 'records must be an array' });
      }
      const lockResult = withLock(function() {
        const results = data.records.map(r => processOpremaRecord(r, data.kooperantID));
        return {
          success: true,
          processed: results.length,
          succeeded: results.filter(r => r && r.success).length,
          failed: results.filter(r => !r || !r.success).length,
          results: results
        };
      });
      return jsonResponse(lockResult);
    }

    if (data.action === 'syncTrosak') {
      if (!Array.isArray(data.records)) {
        return jsonResponse({ success: false, error: 'records must be an array' });
      }
      const lockResult = withLock(function() {
        const results = data.records.map(r => processTrosakRecord(r, data.kooperantID));
        return {
          success: true,
          processed: results.length,
          succeeded: results.filter(r => r && r.success).length,
          failed: results.filter(r => !r || !r.success).length,
          results: results
        };
      });
      return jsonResponse(lockResult);
    }

    if (data.action === 'saveOtkupniListPdf') {
      return jsonResponse(generateOtkupniListPdf(data));
    }

    if (data.action === 'uploadPdf') {
      return jsonResponse(uploadPdfToDrive(data));
    }

    if (data.action === 'saveWarRoomDemand') {
      if (tokenData.role !== 'Management') {
        return jsonResponse({ success: false, error: 'Nemate pristup', code: 403 });
      }
      return jsonResponse(saveWarRoomDemand(data));
    }

    if (data.action === 'removeWarRoomDemand') {
      if (tokenData.role !== 'Management') {
        return jsonResponse({ success: false, error: 'Nemate pristup', code: 403 });
      }
      return jsonResponse(removeWarRoomDemand(data));
    }

    if (data.action === 'updateDemandPrimljeno') {
      if (tokenData.role !== 'Management') {
        return jsonResponse({ success: false, error: 'Nemate pristup', code: 403 });
      }
      return jsonResponse(updateDemandPrimljeno(data));
    }

    if (data.action === 'updateKamionStatus') {
      return jsonResponse(updateKamionStatus(data));
    }

    if (data.action === 'saveDispecer') {
      if (tokenData.role !== 'Management') {
        return jsonResponse({ success: false, error: 'Nemate pristup', code: 403 });
      }
      return jsonResponse(saveDispecer(data));
    }

    if (data.action === 'updateDispecer') {
      if (tokenData.role !== 'Management') {
        return jsonResponse({ success: false, error: 'Nemate pristup', code: 403 });
      }
      return jsonResponse(updateDispecer(data));
    }

    if (data.action === 'removeDispecer') {
      if (tokenData.role !== 'Management') {
        return jsonResponse({ success: false, error: 'Nemate pristup', code: 403 });
      }
      return jsonResponse(removeDispecer(data));
    }

    if (data.action === 'saveIzdavanje') {
      if (tokenData.role !== 'Management') {
        return jsonResponse({ success: false, error: 'Nemate pristup', code: 403 });
      }
      return jsonResponse(saveIzdavanje(data));
    }

    if (data.action === 'parseFiskalniImage') {
      return jsonResponse(parseFiskalniImage(data));
    }

    if (data.action === 'parseFiskalni') {
      return jsonResponse(parseFiskalni(data));
    }

    if (data.action === 'saveFiskalni') {
      return jsonResponse(saveFiskalni(data));
    }

    if (data.action === 'saveFiskalniMapiranje') {
      return jsonResponse(saveFiskalniMapiranje(data));
    }
    
    if (data.action === 'createArtikal') {
      if (tokenData.role !== 'Management') {
        return jsonResponse({ success: false, error: 'Nemate pristup', code: 403 });
      }
      return jsonResponse(createArtikal(data));
    }

    return jsonResponse({ success: false, error: 'Unknown action' });
  } catch (err) {
    return jsonResponse({ success: false, error: err.message });
  }
}

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || '';

    if (action === 'ping') {
      return jsonResponse({ success: true, timestamp: new Date().toISOString() });
    }

    if (action === 'getParcelGeo') {
      return jsonResponse(getParcelGeo(e.parameter.parcelaId));
    }

    if (action === 'getParcelMeteo') {
      return jsonResponse(getParcelMeteo(e.parameter.parcelaId));
    }

    if (action === 'getParcelMeteoLatest') {
      return jsonResponse(getParcelMeteoLatest(e.parameter.parcelaId));
    }

    if (action === 'getAllMeteoLatest') {
      return jsonResponse(getAllMeteoLatest());
    }

    return jsonResponse({ success: false, error: 'Unknown action' });
  } catch (err) {
    return jsonResponse({ success: false, error: err.message });
  }
}

// ============================================================
// AUTH
// ============================================================

function authenticateUser(username, pin) {
  try {
    if (!username || !pin) return { success: false, error: 'Username i PIN su obavezni' };
    
    const cache = CacheService.getScriptCache();
    const attemptsKey = 'ATTEMPTS_' + username.toLowerCase();
    const attempts = parseInt(cache.get(attemptsKey) || '0');
    
    if (attempts >= 5) {
      logLoginAttempt(username, '', false, 'BLOCKED');
      return { success: false, error: 'Previše pokušaja. Sačekajte 15 minuta.' };
    }
    
    const folder = DriveApp.getFolderById(MASTER_FOLDER_ID);
    const files = folder.getFilesByName('Stammdaten');
    if (!files.hasNext()) return { success: false, error: 'System error' };
    
    const ss = SpreadsheetApp.open(files.next());
    const sheet = ss.getSheetByName('Users');
    if (!sheet) return { success: false, error: 'System error' };
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: false, error: 'System error' };
    
    const headers = data[0];
    const colUser = headers.indexOf('Username');
    const colPin = headers.indexOf('PIN');
    const colRole = headers.indexOf('Role');
    const colEntity = headers.indexOf('EntityID');
    const colName = headers.indexOf('DisplayName');
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][colUser]).toLowerCase().trim() === username.toLowerCase().trim()) {
        if (String(data[i][colPin]) === pin) {
          cache.remove(attemptsKey);
          const token = generateToken();
          const entityID = String(data[i][colEntity]);
          const role = String(data[i][colRole]);
          saveToken(token, entityID, role);
          logLoginAttempt(username, entityID, true, 'OK');
          return {
            success: true, token: token, role: role,
            entityID: entityID, displayName: data[i][colName]
          };
        } else {
          cache.put(attemptsKey, String(attempts + 1), 900);
          logLoginAttempt(username, '', false, 'Wrong PIN');
          return { success: false, error: 'Pogrešan PIN' };
        }
      }
    }
    
    cache.put(attemptsKey, String(attempts + 1), 900);
    logLoginAttempt(username, '', false, 'Username not found');
    return { success: false, error: 'Pogrešno korisničko ime ili PIN' };
  } catch (err) {
    return { success: false, error: 'System error' };
  }
}

function generateToken() {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  let token = '';
  for (let i = 0; i < 64; i++) token += chars.charAt(Math.floor(Math.random() * chars.length));
  return token;
}

function saveToken(token, entityID, role) {
  CacheService.getScriptCache().put('TOKEN_' + token,
    JSON.stringify({ entityID: entityID, role: role, created: new Date().toISOString() }), 86400);
}

function validateToken(token) {
  if (!token || token.length < 10) return false;
  return CacheService.getScriptCache().get('TOKEN_' + token) !== null;
}

function getTokenData(token) {
  const d = CacheService.getScriptCache().get('TOKEN_' + token);
  return d ? JSON.parse(d) : null;
}

function logLoginAttempt(username, entityID, success, message) {
  try {
    const folder = DriveApp.getFolderById(MASTER_FOLDER_ID);
    let files = folder.getFilesByName('LoginLog');
    let ss;
    if (files.hasNext()) {
      ss = SpreadsheetApp.open(files.next());
    } else {
      ss = SpreadsheetApp.create('LoginLog');
      const file = DriveApp.getFileById(ss.getId());
      folder.addFile(file);
      DriveApp.getRootFolder().removeFile(file);
      ss.getSheets()[0].getRange(1, 1, 1, 5).setValues([['Timestamp', 'Username', 'EntityID', 'Success', 'Message']]);
    }
    ss.getSheets()[0].appendRow([new Date().toISOString(), username, entityID, success ? 'Da' : 'Ne', message]);
  } catch (e) {}
}

// ============================================================
// OTKUP PROCESSING
// ============================================================

function processRecord(record, otkupacID) {
  try {
    const sheetName = 'OTK-' + (otkupacID || 'UNKNOWN');
    const ss = getOrCreateSheet(sheetName, COLUMNS);
    const sheet = ss.getSheets()[0];

    ensureSheetColumns(sheet, COLUMNS);

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    const idx = headerIndexMap(headers);
    const nowIso = new Date().toISOString();

    if (!record || !record.clientRecordID) {
      return {
        clientRecordID: '',
        success: false,
        error: 'Missing clientRecordID'
      };
    }

    const existingRow = findByColumn(sheet, idx.ClientRecordID, record.clientRecordID);

    // --------------------------------------------------
    // EXISTING RECORD -> idempotent return / light update
    // --------------------------------------------------
    if (existingRow > 0) {
      const existingValues = sheet.getRange(existingRow, 1, 1, sheet.getLastColumn()).getValues()[0];

      const currentServerRecordID = String(getCell(existingValues, idx.ServerRecordID, '') || '');
      const currentVozacID = String(getCell(existingValues, idx.VozacID, '') || '');

      // optional enrichment: fill missing VozacID
      if (typeof idx.VozacID === 'number' && idx.VozacID >= 0) {
        if (record.vozacID && !currentVozacID) {
          sheet.getRange(existingRow, idx.VozacID + 1).setValue(record.vozacID);
        }
      }

      // optional enrichment: keep latest UpdatedAtClient from client
      if (typeof idx.UpdatedAtClient === 'number' && idx.UpdatedAtClient >= 0) {
        if (record.updatedAtClient) {
          sheet.getRange(existingRow, idx.UpdatedAtClient + 1).setValue(record.updatedAtClient);
        }
      }

      // always stamp server-side confirmation time
      if (typeof idx.UpdatedAtServer === 'number' && idx.UpdatedAtServer >= 0) {
        sheet.getRange(existingRow, idx.UpdatedAtServer + 1).setValue(nowIso);
      }

      if (typeof idx.SyncStatus === 'number' && idx.SyncStatus >= 0) {
        sheet.getRange(existingRow, idx.SyncStatus + 1).setValue('Synced');
      }

      if (typeof idx.ReceivedAt === 'number' && idx.ReceivedAt >= 0) {
        sheet.getRange(existingRow, idx.ReceivedAt + 1).setValue(nowIso);
      }

      return {
        clientRecordID: record.clientRecordID,
        success: true,
        status: 'existing',
        serverRecordID: currentServerRecordID,
        updatedAtServer: nowIso,
        row: existingRow
      };
    }

    // --------------------------------------------------
    // NEW RECORD -> insert
    // --------------------------------------------------
    const serverRecordID = generateServerRecordID(otkupacID);

    const rowObj = {
      ClientRecordID: record.clientRecordID || '',
      ServerRecordID: serverRecordID,
      CreatedAtClient: record.createdAtClient || '',
      UpdatedAtClient: record.updatedAtClient || record.createdAtClient || '',
      UpdatedAtServer: nowIso,
      SyncStatus: 'Synced',
      DeviceID: record.deviceID || '',
      OtkupacID: otkupacID || '',
      Datum: record.datum || '',
      KooperantID: record.kooperantID || '',
      KooperantName: record.kooperantName || '',
      VrstaVoca: record.vrstaVoca || '',
      SortaVoca: record.sortaVoca || '',
      Klasa: record.klasa || 'I',
      Kolicina: record.kolicina || 0,
      Cena: record.cena || 0,
      TipAmbalaze: record.tipAmbalaze || '',
      KolAmbalaze: record.kolAmbalaze || 0,
      ParcelaID: record.parcelaID || '',
      VozacID: record.vozacID || '',
      Napomena: record.napomena || '',
      ReceivedAt: nowIso
    };

    const rowValues = headers.map(h => rowObj[h] !== undefined ? rowObj[h] : '');
    sheet.appendRow(rowValues);

    return {
      clientRecordID: record.clientRecordID,
      success: true,
      status: 'inserted',
      serverRecordID: serverRecordID,
      updatedAtServer: nowIso,
      row: sheet.getLastRow()
    };
  } catch (err) {
    return {
      clientRecordID: record && record.clientRecordID ? record.clientRecordID : '',
      success: false,
      error: err.message
    };
  }
}

function uploadPdfToDrive(data) {
  try {
    const folder = DriveApp.getFolderById(MASTER_FOLDER_ID);
    let subFolder;
    const subFolders = folder.getFoldersByName('OtkupniListovi');
    if (subFolders.hasNext()) { subFolder = subFolders.next(); }
    else { subFolder = folder.createFolder('OtkupniListovi'); }

    const bytes = Utilities.base64Decode(data.pdfBase64);
    const blob = Utilities.newBlob(bytes, 'application/pdf', (data.fileName || 'OtkupniList') + '.pdf');
    const file = subFolder.createFile(blob);

    return { success: true, fileId: file.getId(), fileName: file.getName() };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ============================================================
// ZBIRNA PROCESSING
// ============================================================
function processZbirnaRecord(record, vozacID) {
  try {
    const sheetName = 'VOZ-' + (vozacID || 'UNKNOWN');
    const ss = getOrCreateSheet(sheetName, ZBIRNA_COLUMNS);
    const sheet = ss.getSheets()[0];

    ensureSheetColumns(sheet, ZBIRNA_COLUMNS);

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    const idx = headerIndexMap(headers);
    ensurePlainTextColumn(sheet, headers, 'TipAmbalaze');
    ensurePlainTextColumn(sheet, headers, 'BrojZbirne');
    const nowIso = new Date().toISOString();

    if (!record || !record.clientRecordID) {
      return {
        clientRecordID: '',
        success: false,
        error: 'Missing clientRecordID'
      };
    }

    const existingRow = findByColumn(sheet, idx.ClientRecordID, record.clientRecordID);

    // Existing record -> idempotent return / light update
    if (existingRow > 0) {
      const existingValues = sheet.getRange(existingRow, 1, 1, sheet.getLastColumn()).getValues()[0];
      const currentServerRecordID = String(getCell(existingValues, idx.ServerRecordID, '') || '');

      if (typeof idx.UpdatedAtClient === 'number' && idx.UpdatedAtClient >= 0 && record.updatedAtClient) {
        sheet.getRange(existingRow, idx.UpdatedAtClient + 1).setValue(record.updatedAtClient);
      }

      if (typeof idx.UpdatedAtServer === 'number' && idx.UpdatedAtServer >= 0) {
        sheet.getRange(existingRow, idx.UpdatedAtServer + 1).setValue(nowIso);
      }

      if (typeof idx.SyncStatus === 'number' && idx.SyncStatus >= 0) {
        sheet.getRange(existingRow, idx.SyncStatus + 1).setValue('Synced');
      }

      if (typeof idx.ReceivedAt === 'number' && idx.ReceivedAt >= 0) {
        sheet.getRange(existingRow, idx.ReceivedAt + 1).setValue(nowIso);
      }

      return {
        clientRecordID: record.clientRecordID,
        success: true,
        status: 'existing',
        serverRecordID: currentServerRecordID,
        updatedAtServer: nowIso,
        row: existingRow
      };
    }

    // New insert
    const serverRecordID = generateEntityServerID('ZBR', vozacID);

    const rowObj = {
      ClientRecordID: record.clientRecordID || '',
      ServerRecordID: serverRecordID,
      CreatedAtClient: record.createdAtClient || '',
      UpdatedAtClient: record.updatedAtClient || record.createdAtClient || '',
      UpdatedAtServer: nowIso,
      SyncStatus: 'Synced',
      VozacID: vozacID || '',
      Datum: record.datum || '',
      KupacID: record.kupacID || '',
      KupacName: record.kupacName || '',
      VrstaVoca: record.vrstaVoca || '',
      SortaVoca: record.sortaVoca || '',
      KolicinaKlI: record.kolicinaKlI || 0,
      KolicinaKlII: record.kolicinaKlII || 0,
      TipAmbalaze: record.tipAmbalaze || '',
      KolAmbalaze: record.kolAmbalaze || 0,
      Klasa: record.klasa || '',
      OtkupRecordIDs: record.otkupRecordIDs || '',
      ReceivedAt: nowIso,
      BrojZbirne: record.brojZbirne || '',
    };

    const rowValues = headers.map(h => rowObj[h] !== undefined ? rowObj[h] : '');
    sheet.appendRow(rowValues);

    return {
      clientRecordID: record.clientRecordID,
      success: true,
      status: 'inserted',
      serverRecordID: serverRecordID,
      updatedAtServer: nowIso,
      row: sheet.getLastRow()
    };
  } catch (err) {
    return {
      clientRecordID: record && record.clientRecordID ? record.clientRecordID : '',
      success: false,
      error: err.message
    };
  }
}

function getOtkupiForVozac(vozacID) {
  try {
    if (!vozacID) return { success: false, error: 'vozacID required' };
    var registry = getSheetRegistry();
    var otkKeys = Object.keys(registry).filter(function(k) { return k.startsWith('OTK-'); });
    var allRecords = [];

    if (otkKeys.length > 0) {
      for (var i = 0; i < otkKeys.length; i++) {
        try {
          var records = sheetToArray(SpreadsheetApp.openById(registry[otkKeys[i]]).getSheets()[0]);
          records.forEach(function(r) {
            if ((r.VozacID || r.VozaciID || '') === vozacID) {
              r._source = otkKeys[i];
              allRecords.push(r);
            }
          });
        } catch (e) {}
      }
      return { success: true, records: allRecords };
    }

    // fallback
    var folder = DriveApp.getFolderById(MASTER_FOLDER_ID);
    var files = folder.getFiles();
    while (files.hasNext()) {
      var file = files.next();
      var name = file.getName();
      if (name.startsWith('OTK-')) {
        var records2 = sheetToArray(SpreadsheetApp.open(file).getSheets()[0]);
        records2.forEach(function(r) {
          if ((r.VozacID || r.VozaciID || '') === vozacID) {
            r._source = name;
            allRecords.push(r);
          }
        });
      }
    }
    return { success: true, records: allRecords };
  } catch (err) { return { success: false, error: err.message }; }
}

function getZbirneForVozac(vozacID) {
  try {
    if (!vozacID) return { success: false, error: 'vozacID required' };
    const folder = DriveApp.getFolderById(MASTER_FOLDER_ID);
    const files = folder.getFilesByName('VOZ-' + vozacID);
    if (!files.hasNext()) return { success: true, records: [] };
    return { success: true, records: sheetToArray(SpreadsheetApp.open(files.next()).getSheets()[0]) };
  } catch (err) { return { success: false, error: err.message }; }
}

// ============================================================
// AGROMERE PROCESSING
// ============================================================

function processAgromereRecord(record, kooperantID) {
  try {
    const sheetName = 'AGRO-' + (kooperantID || 'UNKNOWN');
    const ss = getOrCreateSheet(sheetName, AGROMERE_COLUMNS);
    const sheet = ss.getSheets()[0];

    ensureSheetColumns(sheet, AGROMERE_COLUMNS);

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    const idx = headerIndexMap(headers);
    const nowIso = new Date().toISOString();

    if (!record || !record.clientRecordID) {
      return { clientRecordID: '', success: false, error: 'Missing clientRecordID' };
    }

    const existingRow = findByColumn(sheet, idx.ClientRecordID, record.clientRecordID);
    if (existingRow > 0) {
      if (typeof idx.SyncStatus === 'number' && idx.SyncStatus >= 0) {
        sheet.getRange(existingRow, idx.SyncStatus + 1).setValue('Synced');
      }
      if (typeof idx.ReceivedAt === 'number' && idx.ReceivedAt >= 0) {
        sheet.getRange(existingRow, idx.ReceivedAt + 1).setValue(nowIso);
      }
      return { clientRecordID: record.clientRecordID, success: true, status: 'existing', row: existingRow };
    }

    const rowObj = {
      ClientRecordID: record.clientRecordID || '',
      CreatedAtClient: record.createdAtClient || '',
      SyncStatus: 'Synced',
      KooperantID: kooperantID || '',
      ParcelaID: record.parcelaID || '',
      Mera: record.mera || '',
      Datum: record.datum || '',
      Vreme: record.vreme || '',
      Napomena: record.napomena || '',
      ReceivedAt: nowIso
    };

    const rowValues = headers.map(h => rowObj[h] !== undefined ? rowObj[h] : '');
    sheet.appendRow(rowValues);

    return { clientRecordID: record.clientRecordID, success: true, status: 'inserted', row: sheet.getLastRow() };
  } catch (err) {
    return { clientRecordID: record && record.clientRecordID ? record.clientRecordID : '', success: false, error: err.message };
  }
}

function getAgromereForKooperant(kooperantID) {
  try {
    if (!kooperantID) return { success: false, error: 'kooperantID required' };
    
    const folder = DriveApp.getFolderById(MASTER_FOLDER_ID);
    const files = folder.getFilesByName('AGRO-' + kooperantID);
    if (!files.hasNext()) return { success: true, records: [] };
    
    const ss = SpreadsheetApp.open(files.next());
    return { success: true, records: sheetToArray(ss.getSheets()[0]) };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ============================================================
// TRETMANI
// ============================================================

function processTretmanRecord(record, kooperantID) {
  try {
    const sheetName = 'TRETMAN-' + (kooperantID || 'UNKNOWN');
    const ss = getOrCreateSheet(sheetName, TRETMAN_COLUMNS);
    const sheet = ss.getSheets()[0];

    ensureSheetColumns(sheet, TRETMAN_COLUMNS);

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    const idx = headerIndexMap(headers);
    const nowIso = new Date().toISOString();

    if (!record || !record.clientRecordID) {
      return {
        clientRecordID: '',
        success: false,
        error: 'Missing clientRecordID'
      };
    }

    const existingRow = findByColumn(sheet, idx.ClientRecordID, record.clientRecordID);

    // Existing record -> idempotent return / light update
    if (existingRow > 0) {
      const existingValues = sheet.getRange(existingRow, 1, 1, sheet.getLastColumn()).getValues()[0];
      const currentServerRecordID = String(getCell(existingValues, idx.ServerRecordID, '') || '');

      if (typeof idx.UpdatedAtClient === 'number' && idx.UpdatedAtClient >= 0 && record.updatedAtClient) {
        sheet.getRange(existingRow, idx.UpdatedAtClient + 1).setValue(record.updatedAtClient);
      }

      if (typeof idx.UpdatedAtServer === 'number' && idx.UpdatedAtServer >= 0) {
        sheet.getRange(existingRow, idx.UpdatedAtServer + 1).setValue(nowIso);
      }

      if (typeof idx.SyncStatus === 'number' && idx.SyncStatus >= 0) {
        sheet.getRange(existingRow, idx.SyncStatus + 1).setValue('Synced');
      }

      if (typeof idx.ReceivedAt === 'number' && idx.ReceivedAt >= 0) {
        sheet.getRange(existingRow, idx.ReceivedAt + 1).setValue(nowIso);
      }

      return {
        clientRecordID: record.clientRecordID,
        success: true,
        status: 'existing',
        serverRecordID: currentServerRecordID,
        updatedAtServer: nowIso,
        row: existingRow
      };
    }

    // New insert
    const serverRecordID = generateEntityServerID('TRT', kooperantID);

    const rowObj = {
      ClientRecordID: record.clientRecordID || '',
      ServerRecordID: serverRecordID,
      CreatedAtClient: record.createdAtClient || '',
      UpdatedAtClient: record.updatedAtClient || record.createdAtClient || '',
      UpdatedAtServer: nowIso,
      SyncStatus: 'Synced',
      KooperantID: kooperantID || '',
      ParcelaID: record.parcelaID || '',
      Datum: record.datum || '',
      Mera: record.mera || '',
      ArtikalID: record.artikalID || '',
      ArtikalNaziv: record.artikalNaziv || '',
      KolicinaUpotrebljena: record.kolicinaUpotrebljena || '',
      JedinicaMere: record.jedinicaMere || '',
      DozaPreporucena: record.dozaPreporucena || '',
      DozaPrimenjena: record.dozaPrimenjena || '',
      OpremaTraktor: record.opremaTraktor || '',
      OpremaPrskalica: record.opremaPrskalica || '',
      OpremaOstalo: record.opremaOstalo || '',
      KarencaDana: record.karencaDana || '',
      DatumBerbeDozvoljeno: record.datumBerbeDozvoljeno || '',
      VremePocetka: record.vremePocetka || '',
      VremeZavrsetka: record.vremeZavrsetka || '',
      TrajanjeMinuta: record.trajanjeMinuta || '',
      GeoLatStart: record.geoLatStart || '',
      GeoLngStart: record.geoLngStart || '',
      GeoLatEnd: record.geoLatEnd || '',
      GeoLngEnd: record.geoLngEnd || '',
      GeoAutoDetect: record.geoAutoDetect || '',
      MeteoTemp: record.meteoTemp || '',
      MeteoWind: record.meteoWind || '',
      MeteoHumidity: record.meteoHumidity || '',
      MeteoOverride: record.meteoOverride || '',
      Napomena: record.napomena || '',
      ReceivedAt: nowIso
    };

    const rowValues = headers.map(h => rowObj[h] !== undefined ? rowObj[h] : '');
    sheet.appendRow(rowValues);

    return {
      clientRecordID: record.clientRecordID,
      success: true,
      status: 'inserted',
      serverRecordID: serverRecordID,
      updatedAtServer: nowIso,
      row: sheet.getLastRow()
    };
  } catch (err) {
    return {
      clientRecordID: record && record.clientRecordID ? record.clientRecordID : '',
      success: false,
      error: err.message
    };
  }
}

function processOpremaRecord(record, kooperantID) {
  try {
    const sheetName = 'OPREMA-' + (kooperantID || 'UNKNOWN');
    const ss = getOrCreateSheet(sheetName, OPREMA_COLUMNS);
    const sheet = ss.getSheets()[0];

    ensureSheetColumns(sheet, OPREMA_COLUMNS);

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    const idx = headerIndexMap(headers);
    const nowIso = new Date().toISOString();

    if (!record || !record.clientRecordID) {
      return {
        clientRecordID: '',
        success: false,
        error: 'Missing clientRecordID'
      };
    }

    const existingRow = findByColumn(sheet, idx.ClientRecordID, record.clientRecordID);

    if (existingRow > 0) {
      const existingValues = sheet.getRange(existingRow, 1, 1, sheet.getLastColumn()).getValues()[0];
      const currentServerRecordID = String(getCell(existingValues, idx.ServerRecordID, '') || '');

      if (typeof idx.UpdatedAtClient === 'number' && idx.UpdatedAtClient >= 0 && record.updatedAtClient) {
        sheet.getRange(existingRow, idx.UpdatedAtClient + 1).setValue(record.updatedAtClient);
      }

      if (typeof idx.UpdatedAtServer === 'number' && idx.UpdatedAtServer >= 0) {
        sheet.getRange(existingRow, idx.UpdatedAtServer + 1).setValue(nowIso);
      }

      if (typeof idx.SyncStatus === 'number' && idx.SyncStatus >= 0) {
        sheet.getRange(existingRow, idx.SyncStatus + 1).setValue('Synced');
      }

      if (typeof idx.ReceivedAt === 'number' && idx.ReceivedAt >= 0) {
        sheet.getRange(existingRow, idx.ReceivedAt + 1).setValue(nowIso);
      }

      return {
        clientRecordID: record.clientRecordID,
        success: true,
        status: 'existing',
        serverRecordID: currentServerRecordID,
        updatedAtServer: nowIso,
        row: existingRow
      };
    }

    const serverRecordID = generateEntityServerID('OPR', kooperantID);

    const rowObj = {
      ClientRecordID: record.clientRecordID || '',
      ServerRecordID: serverRecordID,
      CreatedAtClient: record.createdAtClient || nowIso,
      UpdatedAtClient: record.updatedAtClient || record.createdAtClient || nowIso,
      UpdatedAtServer: nowIso,
      SyncStatus: 'Synced',
      KooperantID: kooperantID || '',
      Naziv: record.naziv || '',
      Tip: record.tip || '',
      ReceivedAt: nowIso
    };

    const rowValues = headers.map(h => rowObj[h] !== undefined ? rowObj[h] : '');
    sheet.appendRow(rowValues);

    return {
      clientRecordID: record.clientRecordID,
      success: true,
      status: 'inserted',
      serverRecordID: serverRecordID,
      updatedAtServer: nowIso,
      row: sheet.getLastRow()
    };
  } catch (err) {
    return {
      clientRecordID: record && record.clientRecordID ? record.clientRecordID : '',
      success: false,
      error: err.message
    };
  }
}

function getTretmaniForKooperant(kooperantID) {
  try {
    if (!kooperantID) return { success: false, error: 'kooperantID required' };
    const folder = DriveApp.getFolderById(MASTER_FOLDER_ID);
    const files = folder.getFilesByName('TRETMAN-' + kooperantID);
    if (!files.hasNext()) return { success: true, records: [] };
    return { success: true, records: sheetToArray(SpreadsheetApp.open(files.next()).getSheets()[0]) };
  } catch (err) { return { success: false, error: err.message }; }
}

function getOpremaForKooperant(kooperantID) {
  try {
    if (!kooperantID) return { success: false, error: 'kooperantID required' };
    const folder = DriveApp.getFolderById(MASTER_FOLDER_ID);
    const files = folder.getFilesByName('OPREMA-' + kooperantID);
    if (!files.hasNext()) return { success: true, records: [] };
    return { success: true, records: sheetToArray(SpreadsheetApp.open(files.next()).getSheets()[0]) };
  } catch (err) { return { success: false, error: err.message }; }
}

function getTroskoviForKooperant(kooperantID) {
  try {
    if (!kooperantID) return { success: false, error: 'kooperantID required' };
    const folder = DriveApp.getFolderById(MASTER_FOLDER_ID);
    const files = folder.getFilesByName('TROSKOVI-' + kooperantID);
    if (!files.hasNext()) return { success: true, records: [] };
    return { success: true, records: sheetToArray(SpreadsheetApp.open(files.next()).getSheets()[0]) };
  } catch (err) { return { success: false, error: err.message }; }
}

// ============================================================
// DATA ENDPOINTS
// ============================================================

function getOtkupiForOtkupac(otkupacID) {
  try {
    if (!otkupacID) return { success: false, error: 'otkupacID required' };
    
    const folder = DriveApp.getFolderById(MASTER_FOLDER_ID);
    const files = folder.getFilesByName('OTK-' + otkupacID);
    if (!files.hasNext()) return { success: true, records: [] };
    
    const ss = SpreadsheetApp.open(files.next());
    return { success: true, records: sheetToArray(ss.getSheets()[0]) };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function getKarticaForKooperant(kooperantID) {
  try {
    if (!kooperantID) return { success: false, error: 'kooperantID required' };
    
    const folder = DriveApp.getFolderById(MASTER_FOLDER_ID);
    const files = folder.getFilesByName('Kartice');
    if (!files.hasNext()) return { success: true, records: [] };
    
    const ss = SpreadsheetApp.open(files.next());
    const sheet = ss.getSheets()[0];
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: true, records: [] };
    
    const headers = data[0];
    const koopCol = headers.indexOf('KooperantID');
    if (koopCol < 0) return { success: true, records: [] };
    
    const records = [];
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][koopCol]) === kooperantID) {
        const obj = {};
        for (let j = 0; j < headers.length; j++) obj[headers[j]] = data[i][j];
        records.push(obj);
      }
    }
    return { success: true, records: records };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ============================================================
// STAMMDATEN
// ============================================================

function getStammdaten() {
    try {
        var folder = DriveApp.getFolderById(MASTER_FOLDER_ID);
        var files = folder.getFilesByName('Stammdaten');
        if (!files.hasNext()) return { success: true, data: {}, message: 'Stammdaten not found' };

        var ss = SpreadsheetApp.open(files.next());
        var result = {};

        var tabs = ['Kooperanti', 'Kulture', 'Config', 'Parcele', 'Stanice', 'Kupci', 'Vozaci', 'Artikli', 'Oprema', 'MagacinKoop'];
        tabs.forEach(function(tab) {
            var sheet = ss.getSheetByName(tab);
            if (sheet) result[tab.toLowerCase()] = sheetToArray(sheet);
        });

        // MeteoLatest
        var meteoSheet = ss.getSheetByName('MeteoLatest');
        if (meteoSheet) result.meteolatest = sheetToArray(meteoSheet);

        // Kartice — proizvodnja za sve kooperante
        // Kartice se NE šalju kroz getStammdaten.
        // Kooperant čita svoju karticu kroz getKartica endpoint.
        // Management čita kroz getMgmtKartica ili getMgmtAll.
        // var karticeFiles = folder.getFilesByName('Kartice');
        // if (karticeFiles.hasNext()) {
        //     var karticeSheet = SpreadsheetApp.open(karticeFiles.next()).getSheets()[0];
        //     if (karticeSheet) result.kartice = sheetToArray(karticeSheet);
        // }

        return { success: true, data: result };
    } catch (err) {
        return { success: false, error: err.message };
    }
}



// ============================================================
// SHEET MANAGEMENT
// ============================================================

function getOrCreateSheet(sheetName, columns) {
  const folder = DriveApp.getFolderById(MASTER_FOLDER_ID);
  const files = folder.getFilesByName(sheetName);
  
  if (files.hasNext()) return SpreadsheetApp.open(files.next());
  
  const ss = SpreadsheetApp.create(sheetName);
  const sheet = ss.getSheets()[0];
  sheet.getRange(1, 1, 1, columns.length).setValues([columns]);
  sheet.getRange(1, 1, 1, columns.length).setFontWeight('bold');
  sheet.setFrozenRows(1);
  
  const file = DriveApp.getFileById(ss.getId());
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);

  // registruj novi sheet u registry
  try {
    var stamFiles = folder.getFilesByName('Stammdaten');
    if (stamFiles.hasNext()) {
      var stamSs = SpreadsheetApp.open(stamFiles.next());
      var regSheet = stamSs.getSheetByName('SheetRegistry');
      if (regSheet) {
        regSheet.appendRow([sheetName, ss.getId(), new Date().toISOString()]);
      }
    }
  } catch (regErr) {}

  return ss;
}

function findByColumn(sheet, colIndex, value) {
  if (!value) return 0;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][colIndex] === value) return i + 1;
  }
  return 0;
}

function sheetToArray(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  const result = [];
  for (let i = 1; i < data.length; i++) {
    const obj = {};
    for (let j = 0; j < headers.length; j++) obj[headers[j]] = data[i][j];
    result.push(obj);
  }
  return result;
}

// ============================================================
// Management
// ============================================================
function getOtkupiByStanica(stanicaID) {
  try {
    if (!stanicaID) return { success: false, error: 'stanicaID required' };
    const folder = DriveApp.getFolderById(MASTER_FOLDER_ID);
    const files = folder.getFilesByName('OTK-' + stanicaID);
    if (!files.hasNext()) return { success: true, records: [] };
    return { success: true, records: sheetToArray(SpreadsheetApp.open(files.next()).getSheets()[0]) };
  } catch (err) { return { success: false, error: err.message }; }
}

function getMgmtReport(tabName) {
  try {
    const folder = DriveApp.getFolderById(MASTER_FOLDER_ID);
    
    let sheet = null;
    let files = folder.getFilesByName('MgmtReports');
    if (files.hasNext()) {
      sheet = SpreadsheetApp.open(files.next()).getSheetByName(tabName);
    }
    if (!sheet) {
      files = folder.getFilesByName('Stammdaten');
      if (files.hasNext()) {
        sheet = SpreadsheetApp.open(files.next()).getSheetByName(tabName);
      }
    }
    if (!sheet) return { success: true, records: [] };
    return { success: true, records: sheetToArray(sheet) };
  } catch (err) { return { success: false, error: err.message }; }
}
function getSaldoOM() { return getMgmtReport('SaldoOM'); }
function getSaldoKupci() { return getMgmtReport('SaldoKupci'); }

function getKarticaAll() {
  try {
    const folder = DriveApp.getFolderById(MASTER_FOLDER_ID);
    const files = folder.getFilesByName('Kartice');
    if (!files.hasNext()) return { records: [] };
    return { records: sheetToArray(SpreadsheetApp.open(files.next()).getSheets()[0]) };
  } catch (err) { return { records: [] }; }
}

function getFaktureByKupac(kupacID) {
  try {
    if (!kupacID) return { success: false, error: 'kupacID required' };
    const folder = DriveApp.getFolderById(MASTER_FOLDER_ID);
    
    let sheet = null;
    let files = folder.getFilesByName('MgmtReports');
    if (files.hasNext()) {
      sheet = SpreadsheetApp.open(files.next()).getSheetByName('Fakture');
    }
    if (!sheet) {
      files = folder.getFilesByName('Stammdaten');
      if (files.hasNext()) {
        sheet = SpreadsheetApp.open(files.next()).getSheetByName('Fakture');
      }
    }
    if (!sheet) return { success: true, records: [] };
    
    const all = sheetToArray(sheet);
    const filtered = all.filter(r => 
      String(r.KupacID) === kupacID || String(r.Kupac) === kupacID
    );
    return { success: true, records: filtered };
  } catch (err) { return { success: false, error: err.message }; }
}

function getFakturaStavke(fakturaID) {
  try {
    if (!fakturaID) return { success: false, error: 'fakturaID required' };
    const folder = DriveApp.getFolderById(MASTER_FOLDER_ID);
    
    let sheet = null;
    let files = folder.getFilesByName('MgmtReports');
    if (files.hasNext()) {
      sheet = SpreadsheetApp.open(files.next()).getSheetByName('FakturaStavke');
    }
    if (!sheet) {
      files = folder.getFilesByName('Stammdaten');
      if (files.hasNext()) {
        sheet = SpreadsheetApp.open(files.next()).getSheetByName('FakturaStavke');
      }
    }
    if (!sheet) return { success: true, records: [] };
    
    const all = sheetToArray(sheet);
    return { success: true, records: all.filter(r => String(r.FakturaID) === fakturaID) };
  } catch (err) { return { success: false, error: err.message }; }
}

function getAllOtkupiSheets() {
  try {
    var registry = getSheetRegistry();
    var allRecords = [];

    // iz registry-ja izvuci sve OTK- ključeve
    var otkKeys = Object.keys(registry).filter(function(k) { return k.startsWith('OTK-'); });

    if (otkKeys.length > 0) {
      for (var i = 0; i < otkKeys.length; i++) {
        try {
          var ss = SpreadsheetApp.openById(registry[otkKeys[i]]);
          var records = sheetToArray(ss.getSheets()[0]);
          records.forEach(function(r) { r._sheetName = otkKeys[i]; });
          allRecords.push.apply(allRecords, records);
        } catch (e) {
          Logger.log('Registry stale for ' + otkKeys[i] + ': ' + e.message);
        }
      }
      return allRecords;
    }

    // fallback: stari folder scan ako registry prazan
    var folder = DriveApp.getFolderById(MASTER_FOLDER_ID);
    var files = folder.getFiles();
    while (files.hasNext()) {
      var file = files.next();
      var name = file.getName();
      if (name.startsWith('OTK-')) {
        var records2 = sheetToArray(SpreadsheetApp.open(file).getSheets()[0]);
        records2.forEach(function(r) { r._sheetName = name; });
        allRecords.push.apply(allRecords, records2);
      }
    }
    return allRecords;
  } catch (err) { return []; }
}

function saveIzdavanje(data) {
  try {
    const folder = DriveApp.getFolderById(MASTER_FOLDER_ID);
    let files = folder.getFilesByName('MgmtReports');
    if (!files.hasNext()) return { success: false, error: 'MgmtReports not found' };
    const ss = SpreadsheetApp.open(files.next());

    let sheet = ss.getSheetByName('Izdavanje');
    if (!sheet) {
      sheet = ss.insertSheet('Izdavanje');
      sheet.getRange(1, 1, 1, IZDAVANJE_COLUMNS.length).setValues([IZDAVANJE_COLUMNS]);
      sheet.getRange(1, 1, 1, IZDAVANJE_COLUMNS.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    const izdavanjeID = 'IZD-' + Date.now();
    const today = Utilities.formatDate(new Date(), 'Europe/Belgrade', 'yyyy-MM-dd');

    sheet.appendRow([
      izdavanjeID,
      today,
      data.kooperantID || '',
      data.kooperantName || '',
      data.parcelaID || '',
      JSON.stringify(data.stavke || []),
      parseFloat(data.ukupnaVrednost) || 0,
      data.izdaoUser || '',
      data.napomena || '',
      '', // SigIzdavalac — popuniće se iz PDF flow-a ako treba
      '', // SigPrimalac
      new Date().toISOString()
    ]);

    return { success: true, izdavanjeID: izdavanjeID };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ============================================================
// PARCEL GEO
// ============================================================

function getParcelGeo(parcelaId) {
  try {
    parcelaId = String(parcelaId || '').trim();
    if (!parcelaId) return { success: false, error: 'Missing parcelaId' };

    const ss = SpreadsheetApp.openById(GEO_SPREADSHEET_ID);
    const sh = ss.getSheetByName(GEO_SHEET_PARCELE);
    if (!sh) return { success: false, error: 'Sheet not found' };

    const values = sh.getDataRange().getValues();
    if (values.length < 2) return { success: false, error: 'No parcel data' };

    const headers = values[0];
    const idx = geoIndexMap(headers);

    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      if (String(row[idx.ParcelaID] || '').trim() === parcelaId) {
        return {
          success: true,
          parcel: {
            ParcelaID: String(row[idx.ParcelaID] || ''),
            KooperantID: String(row[idx.KooperantID] || ''),
            KatBroj: String(row[idx.KatBroj] || ''),
            KatOpstina: String(row[idx.KatOpstina] || ''),
            Kultura: String(row[idx.Kultura] || ''),
            PovrsinaHa: row[idx.PovrsinaHa],
            GeoStatus: String(row[idx.GeoStatus] || ''),
            GeoSource: String(row[idx.GeoSource] || ''),
            Lat: Number(String(row[idx.Lat] || '').replace(',', '.')),
            Lng: Number(String(row[idx.Lng] || '').replace(',', '.')),
            PolygonGeoJSON: String(row[idx.PolygonGeoJSON] || ''),
            MeteoEnabled: String(row[idx.MeteoEnabled] || ''),
            RizikStatus: String(row[idx.RizikStatus] || '')
          }
        };
      }
    }
    return { success: false, error: 'Parcel not found' };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function saveParcelPolygon(body) {
  try {
    const parcelaId = String(body.parcelaId || '').trim();
    const polygonGeoJSON = String(body.polygonGeoJSON || '').trim();
    const lat = body.lat;
    const lng = body.lng;

    if (!parcelaId) return { success: false, error: 'Missing parcelaId' };
    if (!polygonGeoJSON) return { success: false, error: 'Missing polygonGeoJSON' };

    const ss = SpreadsheetApp.openById(GEO_SPREADSHEET_ID);
    const sh = ss.getSheetByName(GEO_SHEET_PARCELE);
    if (!sh) return { success: false, error: 'Sheet not found' };

    const values = sh.getDataRange().getValues();
    if (values.length < 2) return { success: false, error: 'No parcel data' };

    const headers = values[0];
    const idx = geoIndexMap(headers);

    for (let i = 1; i < values.length; i++) {
      if (String(values[i][idx.ParcelaID]).trim() === parcelaId) {
        const rowNumber = i + 1;
        sh.getRange(rowNumber, idx.PolygonGeoJSON + 1).setValue(polygonGeoJSON);
        sh.getRange(rowNumber, idx.Lat + 1).setValue(Number(lat));
        sh.getRange(rowNumber, idx.Lng + 1).setValue(Number(lng));
        sh.getRange(rowNumber, idx.GeoStatus + 1).setValue('polygon');
        sh.getRange(rowNumber, idx.GeoSource + 1).setValue('pwa_draw');
        sh.getRange(rowNumber, idx.MeteoEnabled + 1).setValue('Da');
        sh.getRange(rowNumber, idx.DatumAzuriranja + 1).setValue(new Date());
        if (!values[i][idx.DatumGeoUnosa]) {
          sh.getRange(rowNumber, idx.DatumGeoUnosa + 1).setValue(new Date());
        }
        return { success: true, parcelaId: parcelaId };
      }
    }
    return { success: false, error: 'Parcel not found' };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function geoIndexMap(headers) {
  const map = {};
  headers.forEach((h, i) => { map[String(h).trim()] = i; });
  return map;
}

// ============================================================
// METEO + RISK ASSESSMENT
// ============================================================

// Pragovi po kulturi — default vrednosti
// Override: Config key "MeteoFrost_Visnja" = "-3" itd.
const CROP_THRESHOLDS = {
  'Visnja': { frostWarn: 2, frostDanger: 0, heatWarn: 33, heatDanger: 38, sprayWindMax: 15, sprayRainHours: 6, optimalTempMin: 15, optimalTempMax: 30 },
  'Jabuka': { frostWarn: 2, frostDanger: -1, heatWarn: 32, heatDanger: 37, sprayWindMax: 15, sprayRainHours: 6, optimalTempMin: 12, optimalTempMax: 28 },
  'Sljiva': { frostWarn: 2, frostDanger: -1, heatWarn: 34, heatDanger: 38, sprayWindMax: 15, sprayRainHours: 6, optimalTempMin: 14, optimalTempMax: 30 },
  'Kruska': { frostWarn: 2, frostDanger: -1, heatWarn: 33, heatDanger: 37, sprayWindMax: 15, sprayRainHours: 6, optimalTempMin: 13, optimalTempMax: 28 },
  'Breskva': { frostWarn: 3, frostDanger: 0, heatWarn: 35, heatDanger: 40, sprayWindMax: 12, sprayRainHours: 8, optimalTempMin: 15, optimalTempMax: 32 },
  'Malina': { frostWarn: 2, frostDanger: -1, heatWarn: 30, heatDanger: 35, sprayWindMax: 12, sprayRainHours: 6, optimalTempMin: 12, optimalTempMax: 26 },
  '_default': { frostWarn: 2, frostDanger: 0, heatWarn: 33, heatDanger: 38, sprayWindMax: 15, sprayRainHours: 6, optimalTempMin: 14, optimalTempMax: 30 }
};

function getThresholdsForCrop(kultura) {
  return CROP_THRESHOLDS[kultura] || CROP_THRESHOLDS['_default'];
}

function getParcelMeteo(parcelaId) {
  try {
    parcelaId = String(parcelaId || '').trim();
    if (!parcelaId) return { success: false, error: 'Missing parcelaId' };

    // 1) prvo pokušaj iz MeteoLatest
    const latestResult = getParcelMeteoLatest(parcelaId);
    if (latestResult.success && latestResult.meteo) {
      const m = latestResult.meteo;

      let lastFetchMs = NaN;
      if (m.LastFetch instanceof Date) {
        lastFetchMs = m.LastFetch.getTime();
      } else if (m.LastFetch) {
        lastFetchMs = new Date(String(m.LastFetch)).getTime();
      }

      if (!isNaN(lastFetchMs)) {
        const fetchAge = Date.now() - lastFetchMs;

        // koristi scheduled podatke ako nisu stariji od 12h
        if (fetchAge >= 0 && fetchAge < 12 * 60 * 60 * 1000) {
          const result = {
            success: true,
            parcelaId: parcelaId,
            katBroj: '',
            kultura: m.Kultura || '',
            lat: 0,
            lng: 0,
            fetchedAt: (m.LastFetch instanceof Date)
              ? m.LastFetch.toISOString()
              : String(m.LastFetch || ''),
            source: 'scheduled',
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
              soilMoist_1_3: Number(m.SoilMoist_1_3cm) || 0,
              soilTemp_0: Number(m.SoilTemp_0cm) || 0,
              soilTemp_6: Number(m.SoilTemp_6cm) || 0,
              et0: Number(m.ET0) || 0
            },
            risk: {
              level: m.RiskLevel || 'ok',
              items: Array.isArray(m.RiskItems) ? m.RiskItems : [],
              minTemp48h: Number(m.TempMin48h) || 0,
              maxTemp48h: Number(m.TempMax48h) || 0,
              rain24h: Number(m.Rain24h) || 0
            },
            sprayWindow: Array.isArray(m.SprayWindows) ? m.SprayWindows : [],
            daily: Array.isArray(m.ForecastDaily) ? m.ForecastDaily : []
          };

          const geoResult = getParcelGeo(parcelaId);
          if (geoResult.success && geoResult.parcel) {
            result.katBroj = geoResult.parcel.KatBroj || '';
            result.lat = Number(geoResult.parcel.Lat) || 0;
            result.lng = Number(geoResult.parcel.Lng) || 0;
          }

          return result;
        }
      } else {
        Logger.log('Invalid LastFetch for ' + parcelaId + ': ' + m.LastFetch);
      }
    }

    // 2) fallback na live API
    const ss = SpreadsheetApp.openById(GEO_SPREADSHEET_ID);
    const sh = ss.getSheetByName(GEO_SHEET_PARCELE);
    if (!sh) return { success: false, error: 'Geo sheet not found' };

    const values = sh.getDataRange().getValues();
    if (values.length < 2) return { success: false, error: 'No data' };

    const headers = values[0];
    const idx = geoIndexMap(headers);

    let lat = null, lng = null, kultura = '', katBroj = '';
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][idx.ParcelaID] || '').trim() === parcelaId) {
        lat = Number(String(values[i][idx.Lat] || '').replace(',', '.'));
        lng = Number(String(values[i][idx.Lng] || '').replace(',', '.'));
        kultura = String(values[i][idx.Kultura] || '');
        katBroj = String(values[i][idx.KatBroj] || '');
        break;
      }
    }

    if (!lat || !lng || isNaN(lat) || isNaN(lng)) {
      return { success: false, error: 'Parcela nema geo podatke' };
    }

    const meteoData = fetchOpenMeteo(lat, lng);
    if (!meteoData) return { success: false, error: 'Meteo API error' };

    const thresholds = getThresholdsForCrop(kultura);
    const risk = assessRisk(meteoData, thresholds);
    const sprayWindow = calculateSprayWindow(meteoData, thresholds);

    return {
      success: true,
      parcelaId: parcelaId,
      katBroj: katBroj,
      kultura: kultura,
      lat: lat,
      lng: lng,
      fetchedAt: new Date().toISOString(),
      source: 'live',
      current: meteoData.current,
      daily: meteoData.daily,
      hourly: meteoData.hourly,
      risk: risk,
      sprayWindow: sprayWindow,
      thresholds: thresholds
    };

  } catch (err) {
    return { success: false, error: err.message };
  }
}

function fetchOpenMeteo(lat, lng) {
  try {
    const url = 'https://api.open-meteo.com/v1/forecast?' +
      'latitude=' + lat.toFixed(4) +
      '&longitude=' + lng.toFixed(4) +
      '&current=temperature_2m,relative_humidity_2m,wind_speed_10m,wind_gusts_10m,precipitation,weather_code' +
      '&hourly=temperature_2m,relative_humidity_2m,precipitation_probability,precipitation,wind_speed_10m,dew_point_2m' +
      '&daily=temperature_2m_max,temperature_2m_min,precipitation_sum,precipitation_probability_max,wind_speed_10m_max,sunrise,sunset,uv_index_max,weather_code' +
      '&timezone=Europe/Belgrade' +
      '&forecast_days=7';

    Logger.log('METEO URL: ' + url);
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    Logger.log('METEO STATUS: ' + response.getResponseCode());
    Logger.log('METEO BODY: ' + response.getContentText().substring(0, 300));
    
    if (response.getResponseCode() !== 200) return null;

    const data = JSON.parse(response.getContentText());

    return {
      current: {
        temperature: data.current.temperature_2m,
        humidity: data.current.relative_humidity_2m,
        windSpeed: data.current.wind_speed_10m,
        windGusts: data.current.wind_gusts_10m,
        precipitation: data.current.precipitation,
        weatherCode: data.current.weather_code
      },
      daily: data.daily.time.map((date, i) => ({
        date: date,
        tempMax: data.daily.temperature_2m_max[i],
        tempMin: data.daily.temperature_2m_min[i],
        precipSum: data.daily.precipitation_sum[i],
        precipProbMax: data.daily.precipitation_probability_max[i],
        windMax: data.daily.wind_speed_10m_max[i],
        sunrise: data.daily.sunrise[i],
        sunset: data.daily.sunset[i],
        uvMax: data.daily.uv_index_max[i],
        weatherCode: data.daily.weather_code[i]
      })),
      hourly: {
        time: data.hourly.time,
        temperature: data.hourly.temperature_2m,
        humidity: data.hourly.relative_humidity_2m,
        precipProb: data.hourly.precipitation_probability,
        precip: data.hourly.precipitation,
        wind: data.hourly.wind_speed_10m,
        dewPoint: data.hourly.dew_point_2m
      }
    };
  } catch (err) {
    Logger.log('METEO ERROR: ' + err.message);
    return null;
  }
}
function assessRisk(meteo, thresholds) {
  const risks = [];
  const today = meteo.daily[0] || {};
  const tomorrow = meteo.daily[1] || {};

  // Frost risk — check next 48h hourly
  const now = new Date();
  const in48h = new Date(now.getTime() + 48 * 60 * 60 * 1000);
  let minTemp48h = 99;
  for (let i = 0; i < meteo.hourly.time.length; i++) {
    const t = new Date(meteo.hourly.time[i]);
    if (t >= now && t <= in48h) {
      if (meteo.hourly.temperature[i] < minTemp48h) minTemp48h = meteo.hourly.temperature[i];
    }
  }

  if (minTemp48h <= thresholds.frostDanger) {
    risks.push({ type: 'frost', level: 'danger', message: 'MRAZ! Min ' + minTemp48h.toFixed(1) + '°C u narednih 48h', icon: '🥶' });
  } else if (minTemp48h <= thresholds.frostWarn) {
    risks.push({ type: 'frost', level: 'warning', message: 'Moguc mraz: ' + minTemp48h.toFixed(1) + '°C u narednih 48h', icon: '❄️' });
  }

  // Heat stress
  let maxTemp48h = -99;
  for (let i = 0; i < meteo.hourly.time.length; i++) {
    const t = new Date(meteo.hourly.time[i]);
    if (t >= now && t <= in48h) {
      if (meteo.hourly.temperature[i] > maxTemp48h) maxTemp48h = meteo.hourly.temperature[i];
    }
  }

  if (maxTemp48h >= thresholds.heatDanger) {
    risks.push({ type: 'heat', level: 'danger', message: 'TOPLOTNI STRES! Max ' + maxTemp48h.toFixed(1) + '°C', icon: '🔥' });
  } else if (maxTemp48h >= thresholds.heatWarn) {
    risks.push({ type: 'heat', level: 'warning', message: 'Visoka temperatura: ' + maxTemp48h.toFixed(1) + '°C', icon: '☀️' });
  }

  // Rain risk — next 24h
  let rain24h = 0;
  const in24h = new Date(now.getTime() + 24 * 60 * 60 * 1000);
  for (let i = 0; i < meteo.hourly.time.length; i++) {
    const t = new Date(meteo.hourly.time[i]);
    if (t >= now && t <= in24h) {
      rain24h += meteo.hourly.precip[i] || 0;
    }
  }

  if (rain24h > 20) {
    risks.push({ type: 'rain', level: 'danger', message: 'Obilne padavine: ' + rain24h.toFixed(1) + 'mm/24h', icon: '🌧️' });
  } else if (rain24h > 5) {
    risks.push({ type: 'rain', level: 'warning', message: 'Kisa ocekivana: ' + rain24h.toFixed(1) + 'mm/24h', icon: '🌦️' });
  }

  // Disease risk — humidity + temperature combo (simplified Downy Mildew model)
  const avgHumidity = meteo.current.humidity;
  const currentTemp = meteo.current.temperature;
  if (avgHumidity > 85 && currentTemp > 15 && currentTemp < 28) {
    risks.push({ type: 'disease', level: 'warning', message: 'Povoljni uslovi za bolesti (vlaznost ' + avgHumidity + '%, temp ' + currentTemp.toFixed(1) + '°C)', icon: '🦠' });
  }

  // Overall risk level
  let overallLevel = 'ok';
  if (risks.some(r => r.level === 'danger')) overallLevel = 'danger';
  else if (risks.some(r => r.level === 'warning')) overallLevel = 'warning';

  return {
    level: overallLevel,
    items: risks,
    minTemp48h: minTemp48h,
    maxTemp48h: maxTemp48h,
    rain24h: rain24h
  };
}

function calculateSprayWindow(meteo, thresholds) {
  // Find good spray windows in next 72h
  // Criteria: no rain for N hours, wind < max, no precipitation, temp > 5°C
  const now = new Date();
  const windows = [];
  let currentWindow = null;

  for (let i = 0; i < meteo.hourly.time.length; i++) {
    const t = new Date(meteo.hourly.time[i]);
    if (t < now) continue;
    if (t > new Date(now.getTime() + 72 * 60 * 60 * 1000)) break;

    const temp = meteo.hourly.temperature[i];
    const wind = meteo.hourly.wind[i];
    const precip = meteo.hourly.precip[i] || 0;
    const precipProb = meteo.hourly.precipProb[i] || 0;
    const humidity = meteo.hourly.humidity[i];

    // Check if hour is suitable for spraying
    const suitable = (
      precip < 0.1 &&
      precipProb < 30 &&
      wind < thresholds.sprayWindMax &&
      temp > 5 &&
      temp < 35 &&
      humidity < 90
    );

    // Check dry hours ahead
    let dryAhead = 0;
    if (suitable) {
      for (let j = i + 1; j < Math.min(i + thresholds.sprayRainHours, meteo.hourly.time.length); j++) {
        if ((meteo.hourly.precip[j] || 0) < 0.1) dryAhead++;
        else break;
      }
    }

    const goodWindow = suitable && dryAhead >= (thresholds.sprayRainHours - 1);

    if (goodWindow) {
      if (!currentWindow) {
        currentWindow = {
          start: meteo.hourly.time[i],
          end: meteo.hourly.time[i],
          avgTemp: temp,
          avgWind: wind,
          avgHumidity: humidity,
          hours: 1
        };
      } else {
        currentWindow.end = meteo.hourly.time[i];
        currentWindow.avgTemp = (currentWindow.avgTemp * currentWindow.hours + temp) / (currentWindow.hours + 1);
        currentWindow.avgWind = (currentWindow.avgWind * currentWindow.hours + wind) / (currentWindow.hours + 1);
        currentWindow.avgHumidity = (currentWindow.avgHumidity * currentWindow.hours + humidity) / (currentWindow.hours + 1);
        currentWindow.hours++;
      }
    } else {
      if (currentWindow && currentWindow.hours >= 2) {
        windows.push(currentWindow);
      }
      currentWindow = null;
    }
  }

  // Don't forget last window
  if (currentWindow && currentWindow.hours >= 2) {
    windows.push(currentWindow);
  }

  // Return best 3 windows
  return windows.slice(0, 3).map(w => ({
    start: w.start,
    end: w.end,
    hours: w.hours,
    avgTemp: Math.round(w.avgTemp * 10) / 10,
    avgWind: Math.round(w.avgWind * 10) / 10,
    avgHumidity: Math.round(w.avgHumidity)
  }));
}

function weatherCodeToText(code) {
  const codes = {
    0: 'Vedro', 1: 'Pretezno vedro', 2: 'Delimicno oblacno', 3: 'Oblacno',
    45: 'Magla', 48: 'Magla sa mrazom',
    51: 'Slaba kisa', 53: 'Umerena kisa', 55: 'Jaka kisa',
    61: 'Slaba kisa', 63: 'Umerena kisa', 65: 'Jaka kisa',
    71: 'Slab sneg', 73: 'Umeren sneg', 75: 'Jak sneg',
    80: 'Pljuskovi', 81: 'Umereni pljuskovi', 82: 'Jaki pljuskovi',
    95: 'Grmljavina', 96: 'Grmljavina sa gradom', 99: 'Jaka grmljavina'
  };
  return codes[code] || 'Nepoznato';
}

function weatherCodeToIcon(code) {
  if (code === 0) return '☀️';
  if (code <= 3) return '⛅';
  if (code <= 48) return '🌫️';
  if (code <= 55) return '🌧️';
  if (code <= 65) return '🌧️';
  if (code <= 75) return '❄️';
  if (code <= 82) return '🌦️';
  if (code >= 95) return '⛈️';
  return '🌤️';
}

function getParcelMeteoLatest(parcelaId) {
  try {
    parcelaId = String(parcelaId || '').trim();
    if (!parcelaId) return { success: false, error: 'Missing parcelaId' };

    const ss = SpreadsheetApp.openById(GEO_SPREADSHEET_ID);
    const sheet = ss.getSheetByName('MeteoLatest');
    if (!sheet) return { success: false, error: 'No MeteoLatest sheet' };

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: false, error: 'No data' };

    const headers = data[0];

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === parcelaId) {
        const obj = {};
        for (let j = 0; j < headers.length; j++) {
          obj[String(headers[j]).trim()] = data[i][j];
        }

        try {
          obj.RiskItems = JSON.parse(
            obj['RiskItems (JSON)'] || obj.RiskItems || '[]'
          );
        } catch (e) {
          obj.RiskItems = [];
        }

        try {
          obj.SprayWindows = JSON.parse(
            obj['SprayWindows (JSON)'] || obj.SprayWindows || '[]'
          );
        } catch (e) {
          obj.SprayWindows = [];
        }

        try {
          obj.ForecastDaily = JSON.parse(
            obj['ForecastDaily (JSON)'] || obj.ForecastDaily || '[]'
          );
        } catch (e) {
          obj.ForecastDaily = [];
        }

        return { success: true, meteo: obj };
      }
    }

    return { success: false, error: 'Parcel not found' };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function getAllMeteoLatest() {
  try {
    const ss = SpreadsheetApp.openById(GEO_SPREADSHEET_ID);
    const sheet = ss.getSheetByName('MeteoLatest');
    if (!sheet) return { success: true, records: [] };

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: true, records: [] };

    const headers = data[0];
    const records = [];

    for (let i = 1; i < data.length; i++) {
      const obj = {};
      for (let j = 0; j < headers.length; j++) {
        obj[String(headers[j]).trim()] = data[i][j];
      }

      try {
        obj.RiskItems = JSON.parse(
          obj['RiskItems (JSON)'] || obj.RiskItems || '[]'
        );
      } catch (e) {
        obj.RiskItems = [];
      }

      try {
        obj.SprayWindows = JSON.parse(
          obj['SprayWindows (JSON)'] || obj.SprayWindows || '[]'
        );
      } catch (e) {
        obj.SprayWindows = [];
      }

      try {
        obj.ForecastDaily = JSON.parse(
          obj['ForecastDaily (JSON)'] || obj.ForecastDaily || '[]'
        );
      } catch (e) {
        obj.ForecastDaily = [];
      }

      records.push(obj);
    }

    return { success: true, records: records };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ============================================================
// METEO SCHEDULED FETCH (Time Trigger 4x daily)
// ============================================================

function scheduledMeteoFetch() {
    try {
        var ss = SpreadsheetApp.openById(GEO_SPREADSHEET_ID);
        var parcelSheet = ss.getSheetByName(GEO_SHEET_PARCELE);
        if (!parcelSheet) { Logger.log('No Parcele sheet'); return; }

        var values = parcelSheet.getDataRange().getValues();
        if (values.length < 2) return;

        var headers = values[0];
        var idx = geoIndexMap(headers);

        var locations = {};

        for (var i = 1; i < values.length; i++) {
            var meteoEnabled = String(values[i][idx.MeteoEnabled] || '');
            if (meteoEnabled !== 'Da') continue;

            var lat = Number(String(values[i][idx.Lat] || '').replace(',', '.'));
            var lng = Number(String(values[i][idx.Lng] || '').replace(',', '.'));
            if (!lat || !lng || isNaN(lat) || isNaN(lng)) continue;

            var gridLat = Math.round(lat * 100) / 100;
            var gridLng = Math.round(lng * 100) / 100;
            var key = gridLat.toFixed(2) + ',' + gridLng.toFixed(2);

            if (!locations[key]) {
                locations[key] = { lat: gridLat, lng: gridLng, parcelaIds: [], kulturas: [] };
            }

            locations[key].parcelaIds.push(String(values[i][idx.ParcelaID] || ''));
            locations[key].kulturas.push(String(values[i][idx.Kultura] || ''));
        }

        var locationKeys = Object.keys(locations);
        var locationList = locationKeys.map(function(k) { return locations[k]; });
        Logger.log('Fetching meteo for ' + locationList.length + ' locations (batch)');

        if (locationList.length === 0) return;

        // JEDAN API poziv za sve lokacije
        var batchResults = fetchFullMeteoBatch(locationList);

        if (!batchResults) {
            Logger.log('Batch fetch failed, trying individual...');
            // Fallback na pojedinačne pozive sa retry
            batchResults = [];
            for (var li = 0; li < locationList.length; li++) {
                var meteo = null;
                for (var attempt = 0; attempt < 3; attempt++) {
                    meteo = fetchFullMeteo(locationList[li].lat, locationList[li].lng);
                    if (meteo) break;
                    Logger.log('Retry ' + (attempt + 1) + ' for ' + locationKeys[li]);
                    Utilities.sleep(2000 + attempt * 3000);
                }
                batchResults.push(meteo);
                Utilities.sleep(500);
            }
        }

        // Prepare sheets
        var historySheet = ss.getSheetByName('MeteoHistory');
        if (!historySheet) {
            historySheet = ss.insertSheet('MeteoHistory');
            historySheet.getRange(1, 1, 1, 31).setValues([[
                'Timestamp', 'LatLngKey', 'Lat', 'Lng',
                'TempCurrent', 'TempMin24h', 'TempMax24h', 'TempFeelCurrent',
                'Humidity', 'DewPoint', 'Pressure',
                'PrecipCurrent', 'PrecipSum24h', 'PrecipProbMax24h',
                'WindSpeed', 'WindGusts', 'WindDir',
                'CloudCover', 'Visibility',
                'UVIndex', 'SolarRadiation', 'ET0',
                'SoilMoist_0_1cm', 'SoilMoist_1_3cm', 'SoilMoist_3_9cm', 'SoilMoist_9_27cm',
                'SoilTemp_0cm', 'SoilTemp_6cm', 'SoilTemp_18cm', 'SoilTemp_54cm',
                'WeatherCode'
            ]]);
            historySheet.setFrozenRows(1);
        }

        var latestSheet = ss.getSheetByName('MeteoLatest');
        if (!latestSheet) {
            latestSheet = ss.insertSheet('MeteoLatest');
            latestSheet.getRange(1, 1, 1, 28).setValues([[
                'ParcelaID', 'LastFetch', 'Kultura',
                'Temp', 'TempFeel', 'Humidity', 'DewPoint',
                'Precip', 'PrecipProb', 'Wind', 'WindGusts', 'WindDir',
                'CloudCover', 'UVIndex', 'SolarRadiation',
                'SoilMoist_0_1cm', 'SoilMoist_1_3cm',
                'SoilTemp_0cm', 'SoilTemp_6cm',
                'ET0', 'WeatherCode',
                'TempMin48h', 'TempMax48h', 'Rain24h',
                'RiskLevel', 'RiskItems',
                'SprayWindows',
                'ForecastDaily'
            ]]);
            latestSheet.setFrozenRows(1);
        }

        var now = new Date().toISOString();
        var historyRows = [];
        var latestRows = [];

        for (var ki = 0; ki < locationKeys.length; ki++) {
            var key = locationKeys[ki];
            var loc = locations[key];
            var meteo = batchResults[ki];

            if (!meteo) {
                Logger.log('No data for ' + key);
                continue;
            }

            historyRows.push([
                now, key, loc.lat, loc.lng,
                meteo.current.temp, meteo.daily.tempMin, meteo.daily.tempMax, meteo.current.tempFeel,
                meteo.current.humidity, meteo.current.dewPoint, meteo.current.pressure,
                meteo.current.precip, meteo.daily.precipSum, meteo.daily.precipProbMax,
                meteo.current.windSpeed, meteo.current.windGusts, meteo.current.windDir,
                meteo.current.cloudCover, meteo.current.visibility,
                meteo.current.uvIndex, meteo.current.solarRadiation, meteo.daily.et0,
                meteo.current.soilMoist[0], meteo.current.soilMoist[1], meteo.current.soilMoist[2], meteo.current.soilMoist[3],
                meteo.current.soilTemp[0], meteo.current.soilTemp[1], meteo.current.soilTemp[2], meteo.current.soilTemp[3],
                meteo.current.weatherCode
            ]);

            for (var p = 0; p < loc.parcelaIds.length; p++) {
                var pid = loc.parcelaIds[p];
                var kultura = loc.kulturas[p];
                var thresholds = getThresholdsForCrop(kultura);
                var risk = assessRisk(meteo.forRisk, thresholds);
                var spray = calculateSprayWindow(meteo.forRisk, thresholds);

                var forecastDaily = (meteo.forRisk.daily || []).map(function(d) {
                    return {
                        date: d.date, tempMax: d.tempMax, tempMin: d.tempMin,
                        precipSum: d.precipSum, precipProbMax: d.precipProbMax,
                        windMax: d.windMax, weatherCode: d.weatherCode
                    };
                });

                latestRows.push([
                    pid, now, kultura,
                    meteo.current.temp, meteo.current.tempFeel, meteo.current.humidity, meteo.current.dewPoint,
                    meteo.current.precip, meteo.daily.precipProbMax, meteo.current.windSpeed, meteo.current.windGusts, meteo.current.windDir,
                    meteo.current.cloudCover, meteo.current.uvIndex, meteo.current.solarRadiation,
                    meteo.current.soilMoist[0], meteo.current.soilMoist[1],
                    meteo.current.soilTemp[0], meteo.current.soilTemp[1],
                    meteo.daily.et0, meteo.current.weatherCode,
                    risk.minTemp48h, risk.maxTemp48h, risk.rain24h,
                    risk.level, JSON.stringify(risk.items),
                    JSON.stringify(spray),
                    JSON.stringify(forecastDaily)
                ]);
            }
        }

        if (historyRows.length > 0) {
            historySheet.getRange(historySheet.getLastRow() + 1, 1, historyRows.length, historyRows[0].length).setValues(historyRows);
        }

        if (latestSheet.getLastRow() > 1) {
            latestSheet.getRange(2, 1, latestSheet.getLastRow() - 1, latestSheet.getLastColumn()).clearContent();
        }

        if (latestRows.length > 0) {
            latestSheet.getRange(2, 1, latestRows.length, latestRows[0].length).setValues(latestRows);
        }

        Logger.log('Meteo fetch complete: ' + historyRows.length + ' locations, ' + latestRows.length + ' parcels updated');
    } catch (err) {
        Logger.log('scheduledMeteoFetch error: ' + err.message);
    }
}

function fetchFullMeteoBatch(locations) {
    // Open-Meteo podržava comma-separated lat/lng za do 100 lokacija
    var lats = locations.map(function(l) { return l.lat.toFixed(4); }).join(',');
    var lngs = locations.map(function(l) { return l.lng.toFixed(4); }).join(',');

    var url = 'https://api.open-meteo.com/v1/forecast?' +
        'latitude=' + lats +
        '&longitude=' + lngs +
        '&current=temperature_2m,apparent_temperature,relative_humidity_2m,dew_point_2m,' +
        'precipitation,weather_code,cloud_cover,pressure_msl,wind_speed_10m,wind_direction_10m,' +
        'wind_gusts_10m,visibility,uv_index,shortwave_radiation,' +
        'soil_temperature_0cm,soil_temperature_6cm,soil_temperature_18cm,soil_temperature_54cm,' +
        'soil_moisture_0_to_1cm,soil_moisture_1_to_3cm,soil_moisture_3_to_9cm,soil_moisture_9_to_27cm' +
        '&hourly=temperature_2m,relative_humidity_2m,precipitation_probability,precipitation,' +
        'wind_speed_10m,dew_point_2m' +
        '&daily=temperature_2m_max,temperature_2m_min,precipitation_sum,precipitation_probability_max,' +
        'wind_speed_10m_max,sunrise,sunset,uv_index_max,weather_code,et0_fao_evapotranspiration' +
        '&timezone=Europe/Belgrade' +
        '&forecast_days=7';

    try {
        var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
        var code = response.getResponseCode();
        Logger.log('Batch meteo HTTP ' + code + ' for ' + locations.length + ' locations');

        if (code !== 200) {
            Logger.log('Response: ' + response.getContentText().substring(0, 300));
            return null;
        }

        var data = JSON.parse(response.getContentText());

        // Ako je 1 lokacija, API vraća objekat; ako više — niz
        if (locations.length === 1) {
            return [parseSingleMeteoResponse(data)];
        }

        // Za batch: svaki field je niz nizova
        var results = [];
        for (var i = 0; i < locations.length; i++) {
            results.push(parseBatchMeteoItem(data, i));
        }
        return results;
    } catch (err) {
        Logger.log('fetchFullMeteoBatch error: ' + err.message);
        return null;
    }
}

function parseSingleMeteoResponse(d) {
    return {
        current: {
            temp: d.current.temperature_2m,
            tempFeel: d.current.apparent_temperature,
            humidity: d.current.relative_humidity_2m,
            dewPoint: d.current.dew_point_2m,
            precip: d.current.precipitation,
            weatherCode: d.current.weather_code,
            cloudCover: d.current.cloud_cover,
            pressure: d.current.pressure_msl,
            windSpeed: d.current.wind_speed_10m,
            windDir: d.current.wind_direction_10m,
            windGusts: d.current.wind_gusts_10m,
            visibility: d.current.visibility,
            uvIndex: d.current.uv_index,
            solarRadiation: d.current.shortwave_radiation,
            soilTemp: [
                d.current.soil_temperature_0cm,
                d.current.soil_temperature_6cm,
                d.current.soil_temperature_18cm,
                d.current.soil_temperature_54cm
            ],
            soilMoist: [
                d.current.soil_moisture_0_to_1cm,
                d.current.soil_moisture_1_to_3cm,
                d.current.soil_moisture_3_to_9cm,
                d.current.soil_moisture_9_to_27cm
            ]
        },
        daily: {
            tempMax: d.daily.temperature_2m_max[0],
            tempMin: d.daily.temperature_2m_min[0],
            precipSum: d.daily.precipitation_sum[0],
            precipProbMax: d.daily.precipitation_probability_max[0],
            windMax: d.daily.wind_speed_10m_max[0],
            et0: d.daily.et0_fao_evapotranspiration[0],
            uvMax: d.daily.uv_index_max[0]
        },
        forRisk: {
            current: {
                temperature: d.current.temperature_2m,
                humidity: d.current.relative_humidity_2m,
                windSpeed: d.current.wind_speed_10m,
                windGusts: d.current.wind_gusts_10m,
                precipitation: d.current.precipitation,
                weatherCode: d.current.weather_code
            },
            daily: d.daily.time.map(function(date, j) {
                return {
                    date: date,
                    tempMax: d.daily.temperature_2m_max[j],
                    tempMin: d.daily.temperature_2m_min[j],
                    precipSum: d.daily.precipitation_sum[j],
                    precipProbMax: d.daily.precipitation_probability_max[j],
                    windMax: d.daily.wind_speed_10m_max[j],
                    weatherCode: d.daily.weather_code[j]
                };
            }),
            hourly: {
                time: d.hourly.time,
                temperature: d.hourly.temperature_2m,
                humidity: d.hourly.relative_humidity_2m,
                precipProb: d.hourly.precipitation_probability,
                precip: d.hourly.precipitation,
                wind: d.hourly.wind_speed_10m,
                dewPoint: d.hourly.dew_point_2m
            }
        }
    };
}

function parseBatchMeteoItem(data, idx) {
    // Batch response: data[0].current, data[1].current itd.
    var d = data[idx];
    return parseSingleMeteoResponse(d);
}

function fetchFullMeteo(lat, lng) {
    try {
        const url = 'https://api.open-meteo.com/v1/forecast?' +
            'latitude=' + lat.toFixed(4) +
            '&longitude=' + lng.toFixed(4) +
            '&current=temperature_2m,apparent_temperature,relative_humidity_2m,dew_point_2m,' +
            'precipitation,weather_code,cloud_cover,pressure_msl,wind_speed_10m,wind_direction_10m,' +
            'wind_gusts_10m,visibility,uv_index,shortwave_radiation,' +
            'soil_temperature_0cm,soil_temperature_6cm,soil_temperature_18cm,soil_temperature_54cm,' +
            'soil_moisture_0_to_1cm,soil_moisture_1_to_3cm,soil_moisture_3_to_9cm,soil_moisture_9_to_27cm' +
            '&hourly=temperature_2m,relative_humidity_2m,precipitation_probability,precipitation,' +
            'wind_speed_10m,dew_point_2m' +
            '&daily=temperature_2m_max,temperature_2m_min,precipitation_sum,precipitation_probability_max,' +
            'wind_speed_10m_max,sunrise,sunset,uv_index_max,weather_code,et0_fao_evapotranspiration' +
            '&timezone=Europe/Belgrade' +
            '&forecast_days=7';

        const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
        const code = response.getResponseCode();
        const body = response.getContentText();

        // DODAJ OVO — da vidiš šta server vraća
        Logger.log('fetchFullMeteo ' + lat.toFixed(4) + ',' + lng.toFixed(4) + ' → HTTP ' + code);
        if (code !== 200) {
            Logger.log('Response body: ' + body.substring(0, 300));
            return null;
        }

        const d = JSON.parse(body);

        return {
            current: {
                temp: d.current.temperature_2m,
                tempFeel: d.current.apparent_temperature,
                humidity: d.current.relative_humidity_2m,
                dewPoint: d.current.dew_point_2m,
                precip: d.current.precipitation,
                weatherCode: d.current.weather_code,
                cloudCover: d.current.cloud_cover,
                pressure: d.current.pressure_msl,
                windSpeed: d.current.wind_speed_10m,
                windDir: d.current.wind_direction_10m,
                windGusts: d.current.wind_gusts_10m,
                visibility: d.current.visibility,
                uvIndex: d.current.uv_index,
                solarRadiation: d.current.shortwave_radiation,
                soilTemp: [
                    d.current.soil_temperature_0cm,
                    d.current.soil_temperature_6cm,
                    d.current.soil_temperature_18cm,
                    d.current.soil_temperature_54cm
                ],
                soilMoist: [
                    d.current.soil_moisture_0_to_1cm,
                    d.current.soil_moisture_1_to_3cm,
                    d.current.soil_moisture_3_to_9cm,
                    d.current.soil_moisture_9_to_27cm
                ]
            },
            daily: {
                tempMax: d.daily.temperature_2m_max[0],
                tempMin: d.daily.temperature_2m_min[0],
                precipSum: d.daily.precipitation_sum[0],
                precipProbMax: d.daily.precipitation_probability_max[0],
                windMax: d.daily.wind_speed_10m_max[0],
                et0: d.daily.et0_fao_evapotranspiration[0],
                uvMax: d.daily.uv_index_max[0]
            },
            forRisk: {
                current: {
                    temperature: d.current.temperature_2m,
                    humidity: d.current.relative_humidity_2m,
                    windSpeed: d.current.wind_speed_10m,
                    windGusts: d.current.wind_gusts_10m,
                    precipitation: d.current.precipitation,
                    weatherCode: d.current.weather_code
                },
                daily: d.daily.time.map((date, i) => ({
                    date: date,
                    tempMax: d.daily.temperature_2m_max[i],
                    tempMin: d.daily.temperature_2m_min[i],
                    precipSum: d.daily.precipitation_sum[i],
                    precipProbMax: d.daily.precipitation_probability_max[i],
                    windMax: d.daily.wind_speed_10m_max[i],
                    weatherCode: d.daily.weather_code[i]
                })),
                hourly: {
                    time: d.hourly.time,
                    temperature: d.hourly.temperature_2m,
                    humidity: d.hourly.relative_humidity_2m,
                    precipProb: d.hourly.precipitation_probability,
                    precip: d.hourly.precipitation,
                    wind: d.hourly.wind_speed_10m,
                    dewPoint: d.hourly.dew_point_2m
                }
            }
        };
    } catch (err) {
        Logger.log('fetchFullMeteo error for ' + lat + ',' + lng + ': ' + err.message);
        return null;
    }
}

function testMeteo() {
  const result = getParcelMeteo('PAR-00005');
  Logger.log(JSON.stringify(result).substring(0, 500));
}

function setupMeteoTriggers() {
  // Delete existing meteo triggers
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'scheduledMeteoFetch') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Create 4 daily triggers: 06:00, 12:00, 18:00, 00:00
  [6, 12, 18, 0].forEach(hour => {
    ScriptApp.newTrigger('scheduledMeteoFetch')
      .timeBased()
      .atHour(hour)
      .everyDays(1)
      .inTimezone('Europe/Belgrade')
      .create();
  });

  Logger.log('Meteo triggers created: 06:00, 12:00, 18:00, 00:00');
}

// ============================================================
// KNJIGA POLJA
// ============================================================

function getKooperantProizvodnja(kooperantID) {
    try {
        if (!kooperantID) return { success: false, error: 'kooperantID required' };

        var folder = DriveApp.getFolderById(MASTER_FOLDER_ID);
        var files = folder.getFilesByName('Kartice');
        if (!files.hasNext()) return { success: true, records: [] };

        var ss = SpreadsheetApp.open(files.next());
        var sheet = ss.getSheets()[0];
        if (!sheet) return { success: true, records: [] };

        var data = sheet.getDataRange().getValues();
        if (data.length < 2) return { success: true, records: [] };

        var headers = data[0];
        var colKoop = headers.indexOf('KooperantID');
        var colDatum = headers.indexOf('Datum');
        var colBrojDok = headers.indexOf('BrojDok');
        var colParcela = headers.indexOf('BrojParcele');
        var colOpis = headers.indexOf('Opis');
        var colZaduzenje = headers.indexOf('Zaduzenje');

        var records = [];

        for (var i = 1; i < data.length; i++) {
            if (String(data[i][colKoop]).trim() !== kooperantID) continue;

            var opis = String(data[i][colOpis] || '');
            var zaduzenje = parseFloat(data[i][colZaduzenje]) || 0;

            if (zaduzenje <= 0 || !opis.startsWith('Otkup')) continue;
            if (opis === 'UKUPNO') continue;

            var parsed = parseOpisOtkupa(opis);

            records.push({
                Datum: fmtDateGS(data[i][colDatum]),
                BrojDok: data[i][colBrojDok] || '',
                ParcelaID: colParcela >= 0 ? String(data[i][colParcela] || '').trim() : '',
                VrstaVoca: parsed.vrsta,
                Klasa: parsed.klasa,
                Kolicina: parsed.kolicina,
                Cena: parsed.kolicina > 0 ? Math.round(zaduzenje / parsed.kolicina) : 0,
                Vrednost: zaduzenje
            });
        }

        return { success: true, records: records };
    } catch (err) {
        return { success: false, error: err.message };
    }
}

function parseOpisOtkupa(opis) {
    var m = opis.match(/^Otkup\s+(\S+)\s+(I{1,2})\s+([\d.]+)\s*kg/i);
    if (!m) return { vrsta: '', klasa: 'I', kolicina: 0 };
    return {
        vrsta: m[1],
        klasa: m[2],
        kolicina: parseFloat(m[3].replace(/\./g, '')) || 0
    };
}

// ============================================================
// WAR ROOM DATA
// ============================================================

function getWarRoomDemand() {
  try {
    const today = Utilities.formatDate(new Date(), 'Europe/Belgrade', 'yyyy-MM-dd');
    const folder = DriveApp.getFolderById(MASTER_FOLDER_ID);
    const files = folder.getFilesByName('MgmtReports');
    if (!files.hasNext()) return { success: true, records: [] };
    
    const ss = SpreadsheetApp.open(files.next());
    const sheet = ss.getSheetByName('WarRoomDemand');
    if (!sheet) return { success: true, records: [] };
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: true, records: [] };
    
    const headers = data[0];
    const colDatum = headers.indexOf('Datum');
    const records = [];
    
    for (let i = 1; i < data.length; i++) {
      const datum = fmtDateGS(data[i][colDatum]);
      if (datum !== today) continue;
      const obj = {};
      for (let j = 0; j < headers.length; j++) obj[headers[j]] = data[i][j];
      records.push(obj);
    }
    
    return { success: true, records: records };
  } catch (err) {
    return { success: false, error: err.message };
  }
}
// ============================================================
// SAVE / UPDATE / REMOVE
// ============================================================

function saveWarRoomDemand(data) {
  try {
    const folder = DriveApp.getFolderById(MASTER_FOLDER_ID);
    let files = folder.getFilesByName('MgmtReports');
    let ss;
    if (files.hasNext()) {
      ss = SpreadsheetApp.open(files.next());
    } else {
      return { success: false, error: 'MgmtReports not found' };
    }
    
    let sheet = ss.getSheetByName('WarRoomDemand');
    if (!sheet) {
      sheet = ss.insertSheet('WarRoomDemand');
      sheet.getRange(1, 1, 1, WARROOM_DEMAND_COLUMNS.length).setValues([WARROOM_DEMAND_COLUMNS]);
      sheet.getRange(1, 1, 1, WARROOM_DEMAND_COLUMNS.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
    
    const demandID = 'WRD-' + Date.now();
    const today = Utilities.formatDate(new Date(), 'Europe/Belgrade', 'yyyy-MM-dd');
    
    sheet.appendRow([
      demandID,
      today,
      data.kupacID || '',
      data.kupacName || '',
      parseInt(data.kg) || 0,
      data.vrsta || '',
      data.klasa || '',
      0, // Primljeno
      new Date().toISOString()
    ]);
    
    return { success: true, demandID: demandID };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function removeWarRoomDemand(data) {
  try {
    const demandID = String(data.demandID || '').trim();
    if (!demandID) return { success: false, error: 'Missing demandID' };
    
    const folder = DriveApp.getFolderById(MASTER_FOLDER_ID);
    const files = folder.getFilesByName('MgmtReports');
    if (!files.hasNext()) return { success: false, error: 'Not found' };
    
    const ss = SpreadsheetApp.open(files.next());
    const sheet = ss.getSheetByName('WarRoomDemand');
    if (!sheet) return { success: false, error: 'Sheet not found' };
    
    const data2 = sheet.getDataRange().getValues();
    for (let i = 1; i < data2.length; i++) {
      if (String(data2[i][0]) === demandID) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, error: 'Demand not found' };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function updateDemandPrimljeno(data) {
  try {
    const demandID = String(data.demandID || '').trim();
    const primljeno = parseInt(data.primljeno) || 0;
    if (!demandID) return { success: false, error: 'Missing demandID' };
    
    const folder = DriveApp.getFolderById(MASTER_FOLDER_ID);
    const files = folder.getFilesByName('MgmtReports');
    if (!files.hasNext()) return { success: false, error: 'Not found' };
    
    const ss = SpreadsheetApp.open(files.next());
    const sheet = ss.getSheetByName('WarRoomDemand');
    if (!sheet) return { success: false, error: 'Sheet not found' };
    
    const vals = sheet.getDataRange().getValues();
    const headers = vals[0];
    const colPrimljeno = headers.indexOf('Primljeno');
    
    for (let i = 1; i < vals.length; i++) {
      if (String(vals[i][0]) === demandID) {
        sheet.getRange(i + 1, colPrimljeno + 1).setValue(primljeno);
        return { success: true };
      }
    }
    return { success: false, error: 'Not found' };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function updateKamionStatus(data) {
  try {
    const vozacID = String(data.vozacID || '').trim();
    const status = String(data.status || 'slobodan');
    const ruta = String(data.ruta || '');
    if (!vozacID) return { success: false, error: 'Missing vozacID' };
    
    const folder = DriveApp.getFolderById(MASTER_FOLDER_ID);
    const files = folder.getFilesByName('Stammdaten');
    if (!files.hasNext()) return { success: false, error: 'Not found' };
    
    const ss = SpreadsheetApp.open(files.next());
    let sheet = ss.getSheetByName('KamionStatus');
    if (!sheet) {
      sheet = ss.insertSheet('KamionStatus');
      sheet.getRange(1, 1, 1, KAMION_STATUS_COLUMNS.length).setValues([KAMION_STATUS_COLUMNS]);
      sheet.getRange(1, 1, 1, KAMION_STATUS_COLUMNS.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
    
    const vals = sheet.getDataRange().getValues();
    for (let i = 1; i < vals.length; i++) {
      if (String(vals[i][0]) === vozacID) {
        sheet.getRange(i + 1, 2).setValue(status);
        sheet.getRange(i + 1, 3).setValue(ruta);
        sheet.getRange(i + 1, 4).setValue(new Date().toISOString());
        return { success: true };
      }
    }
    // New row
    sheet.appendRow([vozacID, status, ruta, new Date().toISOString()]);
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// Helper: Date formatting server-side
function fmtDateGS(val) {
  if (!val) return '';
  if (val instanceof Date) return Utilities.formatDate(val, 'Europe/Belgrade', 'yyyy-MM-dd');
  const s = String(val);
  if (s.length >= 10) return s.substring(0, 10);
  return s;
}

// Kombinovani endpoint: vraća demand + plans za danas
function getDispecer() {
    try {
        const today = Utilities.formatDate(new Date(), 'Europe/Belgrade', 'yyyy-MM-dd');
        const folder = DriveApp.getFolderById(MASTER_FOLDER_ID);
        const files = folder.getFilesByName('MgmtReports');
        if (!files.hasNext()) return { success: true, demand: [], plans: [] };
 
        const ss = SpreadsheetApp.open(files.next());
 
        // Demand (isti kao getWarRoomDemand)
        const demand = [];
        const demSheet = ss.getSheetByName('WarRoomDemand');
        if (demSheet) {
            const dd = demSheet.getDataRange().getValues();
            if (dd.length >= 2) {
                const dh = dd[0], colD = dh.indexOf('Datum');
                for (let i = 1; i < dd.length; i++) {
                    if (fmtDateGS(dd[i][colD]) === today) {
                        const obj = {};
                        for (let j = 0; j < dh.length; j++) obj[dh[j]] = dd[i][j];
                        demand.push(obj);
                    }
                }
            }
        }
 
        // Plans
        const plans = [];
        const planSheet = ss.getSheetByName('DispecerPlan');
        if (planSheet) {
            const pd = planSheet.getDataRange().getValues();
            if (pd.length >= 2) {
                const ph = pd[0], colPD = ph.indexOf('Datum'), colSt = ph.indexOf('Status');
                for (let i = 1; i < pd.length; i++) {
                    if (fmtDateGS(pd[i][colPD]) === today && pd[i][colSt] !== 'zavrseno') {
                        const obj = {};
                        for (let j = 0; j < ph.length; j++) obj[ph[j]] = pd[i][j];
                        plans.push(obj);
                    }
                }
            }
        }
 
        return { success: true, demand, plans };
    } catch (err) {
        return { success: false, error: err.message };
    }
}
 
function saveDispecer(data) {
    try {
        const folder = DriveApp.getFolderById(MASTER_FOLDER_ID);
        const files = folder.getFilesByName('MgmtReports');
        if (!files.hasNext()) return { success: false, error: 'MgmtReports not found' };
        const ss = SpreadsheetApp.open(files.next());
 
        let sheet = ss.getSheetByName('DispecerPlan');
        if (!sheet) {
            sheet = ss.insertSheet('DispecerPlan');
            sheet.getRange(1, 1, 1, DISPECER_PLAN_COLUMNS.length).setValues([DISPECER_PLAN_COLUMNS]);
            sheet.getRange(1, 1, 1, DISPECER_PLAN_COLUMNS.length).setFontWeight('bold');
            sheet.setFrozenRows(1);
        }
 
        const planID = 'DPL-' + Date.now();
        const today = Utilities.formatDate(new Date(), 'Europe/Belgrade', 'yyyy-MM-dd');
        const now = new Date().toISOString();
 
        sheet.appendRow([
            planID, today,
            data.demandID || '',
            data.vozacID || '', data.vozacName || data.vozacID || '',
            data.stanicaID || '', data.stanicaName || '',
            data.kupacID || '', data.kupacName || '',
            parseInt(data.plannedKg) || 0,
            'planned', now, now
        ]);
 
        return { success: true, planID };
    } catch (err) {
        return { success: false, error: err.message };
    }
}
 
function updateDispecer(data) {
    try {
        const planID = String(data.planID || '').trim();
        const newStatus = String(data.status || '');
        if (!planID) return { success: false, error: 'Missing planID' };
 
        const folder = DriveApp.getFolderById(MASTER_FOLDER_ID);
        const files = folder.getFilesByName('MgmtReports');
        if (!files.hasNext()) return { success: false, error: 'Not found' };
 
        const sheet = SpreadsheetApp.open(files.next()).getSheetByName('DispecerPlan');
        if (!sheet) return { success: false, error: 'Sheet not found' };
 
        const vals = sheet.getDataRange().getValues();
        const headers = vals[0];
        const colStatus = headers.indexOf('Status');
        const colUpdated = headers.indexOf('UpdatedAt');
 
        for (let i = 1; i < vals.length; i++) {
            if (String(vals[i][0]) === planID) {
                sheet.getRange(i + 1, colStatus + 1).setValue(newStatus);
                sheet.getRange(i + 1, colUpdated + 1).setValue(new Date().toISOString());
                return { success: true };
            }
        }
        return { success: false, error: 'Plan not found' };
    } catch (err) {
        return { success: false, error: err.message };
    }
}
 
function removeDispecer(data) {
    try {
        const planID = String(data.planID || '').trim();
        if (!planID) return { success: false, error: 'Missing planID' };
 
        const folder = DriveApp.getFolderById(MASTER_FOLDER_ID);
        const files = folder.getFilesByName('MgmtReports');
        if (!files.hasNext()) return { success: false, error: 'Not found' };
 
        const sheet = SpreadsheetApp.open(files.next()).getSheetByName('DispecerPlan');
        if (!sheet) return { success: false, error: 'Sheet not found' };
 
        const vals = sheet.getDataRange().getValues();
        for (let i = 1; i < vals.length; i++) {
            if (String(vals[i][0]) === planID) {
                sheet.deleteRow(i + 1);
                return { success: true };
            }
        }
        return { success: false, error: 'Not found' };
    } catch (err) {
        return { success: false, error: err.message };
    }
}

function getKamionStatus() {
    try {
        const folder = DriveApp.getFolderById(MASTER_FOLDER_ID);
        const files = folder.getFilesByName('Stammdaten');
        if (!files.hasNext()) return { success: true, records: [] };
        const ss = SpreadsheetApp.open(files.next());
        const sheet = ss.getSheetByName('KamionStatus');
        if (!sheet) return { success: true, records: [] };
        const data = sheet.getDataRange().getValues();
        if (data.length < 2) return { success: true, records: [] };
        const headers = data[0];
        const records = [];
        for (let i = 1; i < data.length; i++) {
            const obj = {};
            for (let j = 0; j < headers.length; j++) obj[headers[j]] = data[i][j];
            records.push(obj);
        }
        return { success: true, records };
    } catch (err) {
        return { success: false, error: err.message };
    }
}

function parseFiskalniImage(data) {
    try {
        var imageBase64 = data.imageBase64;
        var kooperantID = data.kooperantID || '';
        if (!imageBase64) return { success: false, error: 'No image' };

        var bytes = Utilities.base64Decode(imageBase64);
        Logger.log('Image size: ' + bytes.length + ' bytes');

        var blob = Utilities.newBlob(bytes, 'image/jpeg', 'receipt.jpg');

        var resp = UrlFetchApp.fetch('http://api.qrserver.com/v1/read-qr-code/', {
            method: 'post',
            payload: { 'file': blob },
            muteHttpExceptions: true
        });

        var code = resp.getResponseCode();
        var body = resp.getContentText();
        Logger.log('QR service HTTP ' + code);
        Logger.log('QR service response: ' + body.substring(0, 500));

        if (code !== 200) {
            return { success: false, error: 'QR service HTTP ' + code };
        }

        var result = JSON.parse(body);
        var qrData = '';

        if (result && result[0] && result[0].symbol && result[0].symbol[0]) {
            if (result[0].symbol[0].error) {
                return { success: false, error: 'QR nije pronađen na slici' };
            }
            qrData = result[0].symbol[0].data || '';
        }

        if (!qrData) return { success: false, error: 'QR nije pronađen' };

        Logger.log('QR decoded: ' + qrData.substring(0, 100));

        return parseFiskalni({
            verificationUrl: qrData,
            kooperantID: kooperantID
        });
    } catch (err) {
        Logger.log('parseFiskalniImage error: ' + err.message);
        return { success: false, error: err.message };
    }
}

function parseFiskalni(data) {
    try {
        var verificationUrl = String(data.verificationUrl || '').trim();
        var kooperantID = String(data.kooperantID || '').trim();
        if (!verificationUrl) return { success: false, error: 'Missing URL' };

        // Duplikat check
        var folder = DriveApp.getFolderById(MASTER_FOLDER_ID);
        var fisFiles = folder.getFilesByName('FISKALNI-' + kooperantID);
        if (fisFiles.hasNext()) {
            var fisSheet = SpreadsheetApp.open(fisFiles.next()).getSheets()[0];
            var fisData = fisSheet.getDataRange().getValues();
            var headers = fisData[0] || [];
            var colUrl = headers.indexOf('VerificationUrl');
            if (colUrl >= 0) {
                for (var i = 1; i < fisData.length; i++) {
                    if (String(fisData[i][colUrl]).trim() === verificationUrl) {
                        return { success: false, error: 'Ovaj račun je već skeniran', duplicate: true };
                    }
                }
            }
        }

        // Fetch SUF API
        var resp = UrlFetchApp.fetch(verificationUrl, {
            method: 'get',
            followRedirects: true,
            muteHttpExceptions: true,
            headers: { 'Accept': 'application/json' }
        });

        var code = resp.getResponseCode();
        if (code < 200 || code >= 300) return { success: false, error: 'HTTP ' + code };

        var rawData = JSON.parse(resp.getContentText());
        var invoiceResult = rawData.invoiceResult || {};

        var journal = '';
        if (typeof rawData.journal === 'string') journal = rawData.journal;
        else if (rawData.journal != null) journal = JSON.stringify(rawData.journal);

        var meta = extractReceiptMetaOtkup(journal);
        var items = parseItemsFromJournalOtkup(journal);

        // Auto-match
        var mapiranje = loadFiskalniMapiranje();
        var artikli = loadArtikliForMatch();

        var matchedItems = items.map(function(item) {
            var match = autoMatchArtikalServer(item.naziv, artikli, mapiranje);
            return {
                naziv: item.naziv,
                kolicina: item.kolicina,
                jedCena: item.jedCena,
                ukupno: item.ukupno,
                pdvStopa: item.stopa || '',
                artikalID: match ? match.artikalID : '',
                artikalNaziv: match ? match.naziv : '',
                matchConfidence: match ? match.confidence : 'none'
            };
        });

        return {
            success: true,
            invoiceNumber: invoiceResult.invoiceNumber || '',
            company: meta.company,
            date: invoiceResult.sdcTime ? String(invoiceResult.sdcTime).substring(0, 10) : '',
            totalAmount: invoiceResult.totalAmount || 0,
            paymentMethod: meta.paymentMethod,
            verificationUrl: verificationUrl,
            items: matchedItems
        };
    } catch (err) {
        return { success: false, error: err.message };
    }
}

function extractReceiptMetaOtkup(journal) {
    var normalized = journal.replace(/\\r\\n/g, '\n').replace(/\\n/g, '\n').replace(/\\r/g, '\n');
    var lines = normalized.split(/\r\n|\n|\r/).map(function(l) { return l.trim(); }).filter(Boolean);
    var company = '';
    var paymentMethod = '';

    for (var i = 0; i < lines.length; i++) {
        if (/^={5,}/.test(lines[i])) {
            if (i + 1 < lines.length && /^\d+$/.test(lines[i + 1])) {
                if (i + 2 < lines.length) company = lines[i + 2];
            }
            break;
        }
    }

    var paymentKeywords = ['Платна картица', 'Готовина', 'Чек', 'Поклон картица', 'Остало', 'Вирман'];
    for (var j = 0; j < lines.length; j++) {
        for (var k = 0; k < paymentKeywords.length; k++) {
            if (lines[j].startsWith(paymentKeywords[k])) {
                paymentMethod = paymentKeywords[k];
                break;
            }
        }
        if (paymentMethod) break;
    }

    return { company: company, paymentMethod: paymentMethod };
}

function parseItemsFromJournalOtkup(journal) {
    var normalized = journal.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
    var lines = normalized.split('\n').map(function(l) { return l.trim(); }).filter(Boolean);
    if (lines.length <= 3) return parseItemsFlatOtkup(journal);

    var joined = [];
    var i = 0;
    while (i < lines.length) {
        var l = lines[i];
        if (/^[-=]{5,}/.test(l) || (l.includes('Назив') && l.includes('Цена'))) { joined.push(l); i++; continue; }
        if (/^[\d.,]+\s+[\d.,]+\s+[\d.,]+$/.test(l)) { joined.push(l); i++; continue; }
        while (i + 1 < lines.length) {
            var next = lines[i + 1];
            if (/^[\d.,]+\s+[\d.,]+\s+[\d.,]+$/.test(next)) break;
            if (/^[-=]{5,}/.test(next)) break;
            if (next.includes('Назив') && next.includes('Цена')) break;
            if (/^(Укупан износ|УКУПНО|ПФР)/i.test(next)) break;
            i++;
            l = l + ' ' + lines[i];
        }
        joined.push(l.replace(/\(\s+(\S)\s*\)/g, '($1)'));
        i++;
    }
    lines = joined;

    var start = -1;
    for (var j = 0; j < lines.length; j++) {
        if (lines[j].includes('Назив') && lines[j].includes('Цена') && lines[j].includes('Кол.')) { start = j + 1; break; }
    }
    if (start < 0) return parseItemsFlatOtkup(journal);

    var items = [];
    i = start;
    while (i < lines.length) {
        var l2 = lines[i];
        if (/^(Укупан износ|УКУПНО|ПФР|={5,})/i.test(l2)) break;
        if (/^-{5,}/.test(l2)) { i++; continue; }
        var nameLine = l2.match(/^(.+)\s*\(\s*(\S)\s*\)\s*$/);
        if (nameLine && i + 1 < lines.length) {
            var nextL = lines[i + 1].trim();
            var nums = nextL.match(/^([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)$/);
            if (nums) {
                items.push({
                    naziv: nameLine[1].trim(),
                    jedCena: parseSerbNumberOtkup(nums[1]),
                    kolicina: parseSerbNumberOtkup(nums[2]),
                    ukupno: parseSerbNumberOtkup(nums[3]),
                    stopa: nameLine[2].trim()
                });
                i += 2; continue;
            }
        }
        i++;
    }
    return items;
}

function parseItemsFlatOtkup(text) {
    var pattern = /(.+?)\s*\((\S)\)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)/g;
    var items = [];
    var m;
    while ((m = pattern.exec(text)) !== null) {
        var naziv = m[1].trim();
        if (/Назив|Цена|Кол\.|Укупно|Укупан|ПФР|Ознака|Стопа|Порез/.test(naziv)) continue;
        if (naziv.length < 3) continue;
        items.push({
            naziv: naziv,
            jedCena: parseSerbNumberOtkup(m[3]),
            kolicina: parseSerbNumberOtkup(m[4]),
            ukupno: parseSerbNumberOtkup(m[5]),
            stopa: m[2].trim()
        });
    }
    return items;
}

function parseSerbNumberOtkup(s) {
    if (s == null) return 0;
    var str = s.toString().trim();
    var normalized = str.replace(/\./g, '').replace(',', '.');
    var num = Number(normalized);
    return isNaN(num) ? 0 : num;
}

function loadFiskalniMapiranje() {
    try {
        var folder = DriveApp.getFolderById(MASTER_FOLDER_ID);
        var files = folder.getFilesByName('Stammdaten');
        if (!files.hasNext()) return [];
        var ss = SpreadsheetApp.open(files.next());
        var sheet = ss.getSheetByName('FiskalniMapiranje');
        if (!sheet) return [];
        return sheetToArray(sheet);
    } catch (e) { return []; }
}

function loadArtikliForMatch() {
    try {
        var folder = DriveApp.getFolderById(MASTER_FOLDER_ID);
        var files = folder.getFilesByName('Stammdaten');
        if (!files.hasNext()) return [];
        var ss = SpreadsheetApp.open(files.next());
        var sheet = ss.getSheetByName('Artikli');
        if (!sheet) return [];
        return sheetToArray(sheet);
    } catch (e) { return []; }
}

function autoMatchArtikalServer(fiskalniNaziv, artikli, mapiranja) {
    var fn = String(fiskalniNaziv).toUpperCase().trim();

    var mapped = (mapiranja || []).find(function(m) {
        return String(m.FiskalniNaziv || '').toUpperCase().trim() === fn;
    });
    if (mapped) return { artikalID: mapped.ArtikalID, naziv: mapped.ArtikalNaziv, confidence: 'mapped' };

    var exact = (artikli || []).find(function(a) {
        return String(a.Naziv || '').toUpperCase().trim() === fn;
    });
    if (exact) return { artikalID: exact.ArtikalID, naziv: exact.Naziv, confidence: 'exact' };

    var contains = (artikli || []).find(function(a) {
        var an = String(a.Naziv || '').toUpperCase().trim();
        return fn.includes(an) || an.includes(fn);
    });
    if (contains) return { artikalID: contains.ArtikalID, naziv: contains.Naziv, confidence: 'fuzzy' };

    var keywords = fn.split(/[\s,.\-\/]+/).filter(function(w) { return w.length > 2; });
    var bestMatch = null;
    var bestScore = 0;
    for (var i = 0; i < (artikli || []).length; i++) {
        var an = String(artikli[i].Naziv || '').toUpperCase();
        var score = 0;
        keywords.forEach(function(kw) { if (an.includes(kw)) score++; });
        if (score > bestScore && score >= 2) { bestScore = score; bestMatch = artikli[i]; }
    }
    if (bestMatch) return { artikalID: bestMatch.ArtikalID, naziv: bestMatch.Naziv, confidence: 'fuzzy' };

    return null;
}

function saveFiskalni(data) {
    try {
        var kooperantID = String(data.kooperantID || '').trim();
        if (!kooperantID) return { success: false, error: 'Missing kooperantID' };

        var stavke = data.stavke || [];
        if (!stavke.length) return { success: false, error: 'Nema stavki' };

        var sheetName = 'FISKALNI-' + kooperantID;
        var columns = [
            'ClientRecordID', 'CreatedAtClient', 'SyncStatus',
            'KooperantID',
            'InvoiceNumber', 'Company', 'Date', 'VerificationUrl',
            'Naziv', 'ArtikalID', 'ArtikalNaziv',
            'Kolicina', 'JedCena', 'Ukupno',
            'PDVStopa', 'Matched',
            'ReceivedAt'
        ];

        var ss = getOrCreateSheet(sheetName, columns);
        var sheet = ss.getSheets()[0];
        var now = new Date().toISOString();

        var count = 0;
        for (var i = 0; i < stavke.length; i++) {
            var s = stavke[i];
            sheet.appendRow([
                s.clientRecordID || ('FIS-' + Date.now() + '-' + count),
                s.createdAtClient || now,
                'Synced',
                kooperantID,
                data.invoiceNumber || '',
                data.company || '',
                data.date || '',
                data.verificationUrl || '',
                s.naziv || '',
                s.artikalID || '',
                s.artikalNaziv || '',
                s.kolicina || 0,
                s.jedCena || 0,
                s.ukupno || 0,
                s.pdvStopa || '',
                s.artikalID ? 'Da' : 'Ne',
                now
            ]);
            count++;
        }

        return { success: true, saved: count };
    } catch (err) {
        return { success: false, error: err.message };
    }
}

function saveFiskalniMapiranje(data) {
  try {
    var folder = DriveApp.getFolderById(MASTER_FOLDER_ID);
    var files = folder.getFilesByName('Stammdaten');
    if (!files.hasNext()) return { success: false, error: 'Stammdaten not found' };
    var ss = SpreadsheetApp.open(files.next());
    var sheet = ss.getSheetByName('FiskalniMapiranje');
    if (!sheet) {
      sheet = ss.insertSheet('FiskalniMapiranje');
      sheet.getRange(1, 1, 1, FISKALNI_MAP_COLUMNS.length).setValues([FISKALNI_MAP_COLUMNS]);
      sheet.getRange(1, 1, 1, FISKALNI_MAP_COLUMNS.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    var mappings = data.mappings || [];
    if (!mappings.length) return { success: true, saved: 0 };

    var now = new Date().toISOString();
    var count = 0;
    for (var i = 0; i < mappings.length; i++) {
      var m = mappings[i];
      sheet.appendRow([
        m.fiskalniNaziv || '',
        m.artikalID || '',
        m.artikalNaziv || '',
        data.kooperantID || '',
        now
      ]);
      count++;
    }
    return { success: true, saved: count };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function createArtikal(data) {
  try {
    var folder = DriveApp.getFolderById(MASTER_FOLDER_ID);
    var files = folder.getFilesByName('Stammdaten');
    if (!files.hasNext()) return { success: false, error: 'Stammdaten not found' };
    var ss = SpreadsheetApp.open(files.next());
    var sheet = ss.getSheetByName('Artikli');
    if (!sheet) return { success: false, error: 'Artikli sheet not found' };

    var artikalID = 'ART-' + Date.now();
    sheet.appendRow([
      artikalID,
      data.naziv || '',
      data.tip || '',
      data.jedinicaMere || '',
      data.cenaPoJedinici || 0,
      data.dozaPoHa || '',
      data.kultura || '',
      data.pakovanje || '',
      data.barKod || '',
      data.karenca || '',
      'Da'
    ]);

    return { success: true, artikalID: artikalID, naziv: data.naziv };
  } catch (err) {
    return { success: false, error: err.message };
  }
}
// ============================================================
// HELPERS
// ============================================================

function withLock(fn) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); // čekaj do 15s
  } catch (e) {
    return { success: false, error: 'Server zauzet, pokušajte ponovo', code: 503 };
  }
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function getSheetRegistry() {
  try {
    var folder = DriveApp.getFolderById(MASTER_FOLDER_ID);
    var files = folder.getFilesByName('Stammdaten');
    if (!files.hasNext()) return {};
    var ss = SpreadsheetApp.open(files.next());
    var sheet = ss.getSheetByName('SheetRegistry');
    if (!sheet) return {};
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return {};
    var registry = {};
    for (var i = 1; i < data.length; i++) {
      var name = String(data[i][0] || '').trim();
      var fileId = String(data[i][1] || '').trim();
      if (name && fileId) registry[name] = fileId;
    }
    return registry;
  } catch (e) { return {}; }
}

function rebuildSheetRegistry() {
  try {
    var folder = DriveApp.getFolderById(MASTER_FOLDER_ID);
    var stamFiles = folder.getFilesByName('Stammdaten');
    if (!stamFiles.hasNext()) return { success: false, error: 'Stammdaten not found' };
    var ss = SpreadsheetApp.open(stamFiles.next());

    var sheet = ss.getSheetByName('SheetRegistry');
    if (!sheet) {
      sheet = ss.insertSheet('SheetRegistry');
      sheet.getRange(1, 1, 1, 3).setValues([['SheetName', 'FileID', 'UpdatedAt']]);
      sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    // clear existing
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
    }

    var files = folder.getFiles();
    var rows = [];
    var now = new Date().toISOString();
    while (files.hasNext()) {
      var file = files.next();
      var name = file.getName();
      if (name.startsWith('OTK-') || name.startsWith('VOZ-') ||
          name.startsWith('TRETMAN-') || name.startsWith('OPREMA-') ||
          name.startsWith('TROSKOVI-') || name.startsWith('AGRO-') ||
          name.startsWith('FISKALNI-') ||
          name === 'Kartice' || name === 'MgmtReports' || name === 'LoginLog') {
        rows.push([name, file.getId(), now]);
      }
    }

    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, 3).setValues(rows);
    }

    return { success: true, registered: rows.length };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function openSheetByRegistry(sheetName, registry) {
  // pokušaj iz registry-ja
  if (registry && registry[sheetName]) {
    try {
      return SpreadsheetApp.openById(registry[sheetName]);
    } catch (e) {
      // file ID stale — fallback na folder lookup
    }
  }
  // fallback: klasičan folder lookup
  var folder = DriveApp.getFolderById(MASTER_FOLDER_ID);
  var files = folder.getFilesByName(sheetName);
  if (!files.hasNext()) return null;
  return SpreadsheetApp.open(files.next());
}


function headerIndexMap(headers) {
  const map = {};
  headers.forEach((h, i) => {
    map[String(h).trim()] = i;
  });
  return map;
}

function ensureSheetColumns(sheet, requiredColumns) {
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());

  const missing = requiredColumns.filter(col => headers.indexOf(col) === -1);
  if (missing.length === 0) return;

  const nextHeaders = headers.concat(missing);
  sheet.getRange(1, 1, 1, nextHeaders.length).setValues([nextHeaders]);
  sheet.getRange(1, 1, 1, nextHeaders.length).setFontWeight('bold');
  sheet.setFrozenRows(1);
}

function generateServerRecordID(otkupacID) {
  const tz = 'Europe/Belgrade';
  const stamp = Utilities.formatDate(new Date(), tz, 'yyyyMMddHHmmss');
  const rand = Math.floor(Math.random() * 10000);
  return 'OTK-' + String(otkupacID || 'UNK') + '-' + stamp + '-' + rand;
}

function generateEntityServerID(prefix, entityID) {
  const tz = 'Europe/Belgrade';
  const stamp = Utilities.formatDate(new Date(), tz, 'yyyyMMddHHmmss');
  const rand = Math.floor(Math.random() * 10000);
  return String(prefix || 'REC') + '-' + String(entityID || 'UNK') + '-' + stamp + '-' + rand;
}

function getCell(row, idx, fallback) {
  if (typeof idx !== 'number' || idx < 0) return fallback;
  return row[idx];
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function ensurePlainTextColumn(sheet, headers, columnName) {
  const idx = headers.indexOf(columnName);
  if (idx < 0) return;

  const col = idx + 1;
  const rows = Math.max(sheet.getMaxRows() - 1, 1);
  sheet.getRange(2, col, rows, 1).setNumberFormat('@');
}

function quickMeteoDebug() {
    const ss = SpreadsheetApp.openById(GEO_SPREADSHEET_ID);
    const sh = ss.getSheetByName(GEO_SHEET_PARCELE);
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const idx = geoIndexMap(headers);
    
    Logger.log('MeteoEnabled kolona index: ' + idx.MeteoEnabled);
    Logger.log('Lat kolona index: ' + idx.Lat);
    Logger.log('Lng kolona index: ' + idx.Lng);
    Logger.log('Headers: ' + JSON.stringify(headers));
    
    // Probaj fetch za jednu parcelu
    scheduledMeteoFetch();
}
