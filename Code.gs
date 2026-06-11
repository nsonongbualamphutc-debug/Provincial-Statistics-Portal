/**
 * ============================================================
 *  ระบบหลังบ้าน — ข้อมูลสถิติจังหวัดหนองบัวลำภู
 *  Google Apps Script Backend
 * ============================================================
 *  ⚠️ ไฟล์นี้ deploy บน Google Apps Script เท่านั้น
 *  ❌ ห้ามนำขึ้น GitHub (มี hash รหัสผ่านอยู่ภายใน)
 *
 *  วิธีติดตั้ง:
 *  1. สร้าง Google Sheet ใหม่ → Extensions → Apps Script
 *  2. วางโค้ดนี้ลงไปทั้งหมด
 *  3. Deploy → New deployment → Web app
 *     - Execute as: Me
 *     - Who has access: Anyone
 *  4. คัดลอก Web app URL ไปใส่ใน index.html และ admin.html
 *     (ตัวแปร GAS_URL)
 * ============================================================
 */

// ---------- ค่าความปลอดภัย (อยู่ฝั่ง server เท่านั้น) ----------
var PEPPER = 'NBL-STAT-PEPPER-x7Qm2026';
var ADMIN_HASH = '36addb662620a6ffd1a115eee2089510c8e0a98a12566d47f4196d574ace04f0';
var TOKEN_SECRET = 'NBL-TOKEN-SECRET-9kPw2026'; // เปลี่ยนได้ตามต้องการ
var TOKEN_HOURS = 8; // อายุ session token (ชั่วโมง)

var CONFIG_SHEET = 'Config';

// ---------- Utilities ----------
function sha256(str) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, str, Utilities.Charset.UTF_8);
  return raw.map(function(b) {
    var v = (b < 0 ? b + 256 : b).toString(16);
    return v.length === 1 ? '0' + v : v;
  }).join('');
}

function makeToken() {
  var expiry = Date.now() + TOKEN_HOURS * 3600 * 1000;
  var sig = sha256(TOKEN_SECRET + '|' + expiry);
  return expiry + '.' + sig;
}

function verifyToken(token) {
  if (!token) return false;
  var parts = String(token).split('.');
  if (parts.length !== 2) return false;
  var expiry = Number(parts[0]);
  if (!expiry || Date.now() > expiry) return false;
  return sha256(TOKEN_SECRET + '|' + expiry) === parts[1];
}

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CONFIG_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CONFIG_SHEET);
    sh.getRange('A1').setValue('config_json');
    sh.getRange('B1').setValue('updated_at');
  }
  return sh;
}

function loadConfig() {
  var sh = getSheet();
  var val = sh.getRange('A2').getValue();
  if (!val) return null;
  try { return JSON.parse(val); } catch (e) { return null; }
}

function saveConfig(cfg) {
  var sh = getSheet();
  sh.getRange('A2').setValue(JSON.stringify(cfg));
  sh.getRange('B2').setValue(new Date());
}

function jsonOut(obj, callback) {
  var text = JSON.stringify(obj);
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + text + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(text)
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------- GET: อ่าน config (สาธารณะ, รองรับ JSONP) ----------
function doGet(e) {
  var p = (e && e.parameter) || {};
  var callback = p.callback || '';

  if (p.action === 'login') {
    // login ผ่าน JSONP เพื่อเลี่ยง CORS
    var clientHash = String(p.h || '');
    if (clientHash && clientHash === ADMIN_HASH) {
      return jsonOut({ ok: true, token: makeToken() }, callback);
    }
    Utilities.sleep(800); // ชะลอ brute force
    return jsonOut({ ok: false, error: 'รหัสผ่านไม่ถูกต้อง' }, callback);
  }

  var cfg = loadConfig();
  return jsonOut({ ok: true, config: cfg }, callback);
}

// ---------- POST: บันทึก config (ต้องมี token) ----------
function doPost(e) {
  try {
    var body = {};
    if (e.postData && e.postData.contents) {
      try { body = JSON.parse(e.postData.contents); }
      catch (err) { body = e.parameter || {}; }
    } else {
      body = e.parameter || {};
    }

    if (body.action !== 'save') {
      return jsonOut({ ok: false, error: 'unknown action' });
    }
    if (!verifyToken(body.token)) {
      return jsonOut({ ok: false, error: 'session หมดอายุ กรุณาเข้าสู่ระบบใหม่' });
    }

    var cfg = typeof body.config === 'string' ? JSON.parse(body.config) : body.config;
    if (!cfg || typeof cfg !== 'object') {
      return jsonOut({ ok: false, error: 'config ไม่ถูกต้อง' });
    }
    saveConfig(cfg);
    return jsonOut({ ok: true });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}
