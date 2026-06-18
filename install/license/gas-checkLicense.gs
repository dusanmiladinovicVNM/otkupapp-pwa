/**
 * gas-checkLicense.gs  — OPCIONI online kill-switch za AgriX licence (SKICA)
 * ---------------------------------------------------------------------------
 * Deploy kao Web App (Execute as: Me, Access: Anyone).
 * URL upisi u tblSEFConfig: LICENSE_CHECK_URL = https://script.google.com/.../exec
 *
 * Klijent (modLicense.CheckOnlineKillSwitch) salje:
 *     POST { "fingerprint": "<otisak>" }
 * Server vraca POTPISAN verdikt (klijent verifikuje VERDICT kljucem, verdict_pub.cer):
 *     { "status": "ok" | "revoked",
 *       "payload": "<fingerprint>|<status>|<YYYY-MM-DD>",
 *       "sig": "<base64 RSA-SHA256 potpis payload-a>" }
 * Datum u payload-u sluzi i kao freshness (klijent odbija stare/replay verdikte).
 *
 * PODESAVANJE (Project Settings -> Script properties):
 *   VERDICT_PRIV_PEM   = sadrzaj verdict_priv.pem (par sa verdict_pub.cer na stanicama)
 *   REVOKED_LIST       = CSV otisaka koji su opozvani (ili koristi Sheet ispod)
 *
 * NAPOMENA: ovde zivi VERDICT privatni kljuc (server-side secret), NE license
 * kljuc. Tako kompromitacija GAS-a ne omogucava falsifikovanje offline licenci.
 */

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents || '{}');
    var fp = String(body.fingerprint || '').trim();
    if (!fp) return _json({ status: 'error' });

    var status = _isRevoked(fp) ? 'revoked' : 'ok';
    var today = Utilities.formatDate(new Date(), 'Etc/GMT', 'yyyy-MM-dd');
    var payload = fp + '|' + status + '|' + today;

    var pem = PropertiesService.getScriptProperties().getProperty('VERDICT_PRIV_PEM');
    var sigBytes = Utilities.computeRsaSha256Signature(payload, pem);
    var sig = Utilities.base64Encode(sigBytes);

    return _json({ status: status, payload: payload, sig: sig });
  } catch (err) {
    return _json({ status: 'error', message: String(err) });
  }
}

function _isRevoked(fp) {
  // Varijanta A: CSV u Script property.
  var csv = PropertiesService.getScriptProperties().getProperty('REVOKED_LIST') || '';
  var set = csv.split(',').map(function (s) { return s.trim().toUpperCase(); });
  if (set.indexOf(fp.toUpperCase()) >= 0) return true;

  // Varijanta B (opciono): kolona A u tabu "Revoked" nekog Sheet-a.
  // var sh = SpreadsheetApp.openById('<ID>').getSheetByName('Revoked');
  // ...

  return false;
}

function _json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
