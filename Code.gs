/**
 * แผนที่เก็บของ — ตัวเชื่อม Google Sheet
 * วางไฟล์นี้ใน Apps Script ของ Google Sheet (ส่วนขยาย → Apps Script)
 * แล้ว Deploy เป็น Web app (Execute as: Me, Who has access: Anyone)
 */

// ตั้งรหัสลับ แล้วใส่รหัสเดียวกันในหน้าแอป (ปุ่ม "ตั้งค่า Google Sheet")
// ใครที่ไม่มีรหัสนี้จะอ่านหรือแก้ข้อมูลไม่ได้
const KEY = 'เปลี่ยนเป็นรหัสของคุณ';

const PHOTO_FOLDER = 'แผนที่เก็บของ - รูป';

const SHEETS = {
  rooms: ['id', 'name', 'cols', 'rows', 'order'],
  units: ['id', 'roomId', 'parentId', 'code', 'kind', 'x', 'y', 'w', 'h', 'note'],
  items: ['id', 'unitId', 'name', 'qty', 'note', 'photoId', 'photoUrl', 'updatedAt']
};
const NUMBER_FIELDS = { cols: 1, rows: 1, order: 1, x: 1, y: 1, w: 1, h: 1, qty: 1 };

/** กดรันฟังก์ชันนี้ 1 ครั้งในหน้า Apps Script เพื่อสร้างแท็บและอนุญาตสิทธิ์ Google Drive */
function setup() {
  Object.keys(SHEETS).forEach(sheet_);
  photoFolder_();
}

function doGet(e) { return handle_((e && e.parameter) || {}); }

function doPost(e) {
  let p = {};
  try { p = JSON.parse(e.postData.contents); } catch (err) {}
  return handle_(p);
}

function handle_(p) {
  try {
    if (KEY && String(p.key || '') !== KEY) return out_({ ok: false, error: 'รหัส (KEY) ไม่ถูกต้อง' });
    switch (p.action) {
      case 'all':
        return out_({ ok: true, data: readAll_() });
      case 'upsert':
        withLock_(() => upsert_(p.col, p.id, p.data || {}));
        return out_({ ok: true });
      case 'delete':
        withLock_(() => remove_(p.col, p.id));
        return out_({ ok: true });
      case 'upload':
        return out_(Object.assign({ ok: true }, upload_(p)));
      case 'deletePhoto':
        trashPhoto_(p.id);
        return out_({ ok: true });
      default:
        return out_({ ok: true, app: 'storage-map' });
    }
  } catch (err) {
    return out_({ ok: false, error: String((err && err.message) || err) });
  }
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function sheet_(col) {
  const fields = SHEETS[col];
  if (!fields) throw new Error('ไม่รู้จักตาราง ' + col);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(col);
  if (!sh) {
    sh = ss.insertSheet(col);
    sh.getRange(1, 1, sh.getMaxRows(), fields.length).setNumberFormat('@');
    sh.getRange(1, 1, 1, fields.length).setValues([fields]).setFontWeight('bold');
    sh.setFrozenRows(1);
    return sh;
  }
  // เติมหัวคอลัมน์ที่ยังไม่มี (เผื่ออัปเดตแอปในอนาคต)
  const head = headers_(sh);
  const missing = fields.filter(f => head.indexOf(f) < 0);
  if (missing.length) {
    const startCol = head.length + 1;
    sh.getRange(1, startCol, 1, missing.length).setValues([missing]).setFontWeight('bold');
    sh.getRange(1, startCol, sh.getMaxRows(), missing.length).setNumberFormat('@');
  }
  return sh;
}

function headers_(sh) {
  const n = sh.getLastColumn();
  if (!n) return [];
  return sh.getRange(1, 1, 1, n).getValues()[0].map(String).filter(h => h !== '');
}

function readAll_() {
  const data = {};
  Object.keys(SHEETS).forEach(col => {
    const sh = sheet_(col);
    const values = sh.getDataRange().getValues();
    const head = values.shift().map(String);
    data[col] = values
      .filter(row => String(row[0]) !== '')
      .map(row => {
        const o = {};
        head.forEach((h, i) => {
          if (!h) return;
          const v = row[i];
          o[h] = NUMBER_FIELDS[h] ? (v === '' ? 0 : Number(v)) : String(v);
        });
        return o;
      });
  });
  return data;
}

function findRow_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}

function upsert_(col, id, data) {
  if (!id) throw new Error('ไม่มี id');
  const sh = sheet_(col);
  const head = headers_(sh);
  const r = findRow_(sh, id);
  const row = r > 0 ? sh.getRange(r, 1, 1, head.length).getValues()[0] : head.map(() => '');
  head.forEach((h, i) => {
    if (h === 'id') row[i] = String(id);
    else if (Object.prototype.hasOwnProperty.call(data, h)) row[i] = data[h] == null ? '' : String(data[h]);
  });
  if (r > 0) sh.getRange(r, 1, 1, head.length).setValues([row]);
  else sh.appendRow(row);
}

function remove_(col, id) {
  const sh = sheet_(col);
  const r = findRow_(sh, id);
  if (r > 0) sh.deleteRow(r);
}

function photoFolder_() {
  const it = DriveApp.getFoldersByName(PHOTO_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(PHOTO_FOLDER);
}

function upload_(p) {
  const bytes = Utilities.base64Decode(p.data);
  const blob = Utilities.newBlob(bytes, p.mime || 'image/jpeg', p.name || 'photo.jpg');
  const file = photoFolder_().createFile(blob);
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (err) {}
  const id = file.getId();
  return { id: id, url: 'https://drive.google.com/thumbnail?id=' + id + '&sz=w1200' };
}

function trashPhoto_(id) {
  if (!id) return;
  try { DriveApp.getFileById(id).setTrashed(true); } catch (err) {}
}
