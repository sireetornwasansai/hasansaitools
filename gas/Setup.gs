/**
 * HA Round — ตั้งค่าชีต + โฟลเดอร์อัปโหลดจาก URL (รันครั้งเดียวจากหน้าแก้ไข Apps Script)
 * วิธีที่ง่ายที่สุด: เลือกฟังก์ชัน setupAll แล้วกด Run ครั้งเดียว (ชีต + โฟลเดอร์ + สร้างรหัสลับ GAS_SECRET ให้)
 * หรือแยกทำทีละอย่าง: setupSheet (ตั้งชีต/โฟลเดอร์) · showSetupStatus (ตรวจว่าพร้อมใช้งานไหม)
 * แก้ 2 ค่าลิงก์ด้านล่างให้เป็นของคุณ
 *  - บันทึก SHEET_ID / FOLDER_ID (ดึงจาก URL) ลง Script properties → ระบบใช้ชีตนี้และโฟลเดอร์นี้ทุกครั้ง
 *  - สร้างแท็บ Issues, Dates, Hait, Files, Log, Memos, Uploads พร้อมจัดรูปแบบหัวตาราง (ไม่แตะข้อมูลที่มีอยู่)
 *  - ทุกไฟล์ที่อัปโหลดจะถูกบันทึกในแท็บ Uploads (เวลา ผู้อัป ชื่อไฟล์ หมวด ขนาด ลิงก์ไฟล์ ลิงก์โฟลเดอร์) และเก็บในโฟลเดอร์ตาม URL นี้
 */
var SETUP_SHEET_URL = 'https://docs.google.com/spreadsheets/d/1F4jYbvrA9SBc1avqQoF0V89OXCnoCEHmjQG89OmsZ5s/edit';
var SETUP_FOLDER_URL = 'https://drive.google.com/drive/folders/1I5b1BmTl99rPEJ-5dUZselkQJb5w6sbq';

function idFromUrl_(url, kind) {
  var re = kind === 'folder' ? /\/folders\/([A-Za-z0-9_-]{10,})/ : /\/d\/([A-Za-z0-9_-]{10,})/;
  var m = String(url || '').match(re);
  return m ? m[1] : (/^[A-Za-z0-9_-]{20,}$/.test(String(url || '').trim()) ? String(url).trim() : '');   // รับ ID เปล่าๆ ได้ด้วย
}

function setupSheet() {
  var sid = idFromUrl_(SETUP_SHEET_URL, 'sheet'), fid = idFromUrl_(SETUP_FOLDER_URL, 'folder');
  if (!sid) throw new Error('SETUP_SHEET_URL ไม่ใช่ลิงก์ Google Sheet ที่ถูกต้อง');
  if (!fid) throw new Error('SETUP_FOLDER_URL ไม่ใช่ลิงก์โฟลเดอร์ Google Drive ที่ถูกต้อง');
  var ss = SpreadsheetApp.openById(sid);                 // ถ้าไม่มีสิทธิ์/ID ผิด จะหยุดตรงนี้ ยังไม่เขียนค่าใดๆ
  var folder = DriveApp.getFolderById(fid);
  var P = PropertiesService.getScriptProperties();
  P.setProperty('SHEET_ID', sid); P.setProperty('FOLDER_ID', fid);
  P.setProperty('SHEET_URL', ss.getUrl()); P.setProperty('FOLDER_URL', folder.getUrl());

  Object.keys(SHEETS).forEach(function (name) {
    var s = sh_(name), cols = SHEETS[name].length;       // sh_ สร้างแท็บและหัวตารางให้ถ้ายังไม่มี
    var head = s.getRange(1, 1, 1, cols);
    head.setValues([SHEETS[name]]).setFontWeight('bold').setBackground('#1f3a5f').setFontColor('#ffffff');
    s.setFrozenRows(1);
    for (var c = 1; c <= cols; c++) s.setColumnWidth(c, SHEETS[name][c - 1] === 'json' ? 160 : 110);
  });
  ['Sheet1', 'ชีต1'].forEach(function (n) {              // ลบแท็บเปล่าที่ Google สร้างให้ตอนสร้างชีตใหม่
    var d = ss.getSheetByName(n);
    if (d && ss.getSheets().length > 1 && d.getLastRow() <= 1 && d.getLastColumn() <= 1 && !d.getRange(1, 1).getValue()) ss.deleteSheet(d);
  });
  folder.getName();                                      // ยืนยันว่าเข้าถึงโฟลเดอร์ได้
  Logger.log('พร้อมใช้งาน');
  Logger.log('ชีต: ' + ss.getName() + ' → ' + ss.getUrl());
  Logger.log('โฟลเดอร์อัปโหลด: ' + folder.getName() + ' → ' + folder.getUrl());
  Logger.log('แท็บ: ' + Object.keys(SHEETS).join(', '));
}

/** ติดตั้งครั้งเดียวจบ: ชีต/โฟลเดอร์ */
function setupAll() {
  setupSheet();
  Logger.log('— — —');
  Logger.log('ไม่ต้องตั้ง GAS_SECRET แล้ว ใครก็ใช้งานได้ผ่านลิงก์แอป');
  Logger.log('ขั้นต่อไป: Deploy → New deployment → Web app (Execute as: Me, Who has access: Anyone) แล้วนำลิงก์ /exec ไปใส่ GAS_URL ใน Vercel');
  showSetupStatus();
}

/** ตรวจความพร้อมทั้งระบบ — รันแล้วดู Execution log (✔ = พร้อม, ✘ = ต้องแก้) */
function showSetupStatus() {
  var P = PropertiesService.getScriptProperties(), out = [];
  function chk(ok, okMsg, badMsg) { out.push((ok ? '✔ ' : '✘ ') + (ok ? okMsg : badMsg)); }
  try { var ss = openSS_(); chk(true, 'ชีต: ' + ss.getName() + ' (' + Object.keys(SHEETS).filter(function (n) { return ss.getSheetByName(n); }).length + '/' + Object.keys(SHEETS).length + ' แท็บ)', ''); }
  catch (e) { chk(false, '', 'เปิดชีตไม่ได้ → รัน setupSheet (' + e.message + ')'); }
  var fid = P.getProperty('FOLDER_ID');
  if (fid) { try { chk(true, 'โฟลเดอร์ไฟล์แนบ: ' + DriveApp.getFolderById(fid).getName(), ''); } catch (e) { chk(false, '', 'เข้าโฟลเดอร์ไม่ได้ → ตรวจ FOLDER_ID/สิทธิ์'); } }
  else out.push('• โฟลเดอร์ไฟล์แนบ: ยังไม่ตั้ง (ระบบจะสร้าง HA-Round-Files ตอนอัปโหลดครั้งแรก)');
  out.push('• ไม่มีรหัสผ่านผู้ใช้ — ใครมีลิงก์แอปก็ใช้ร่วมกันได้ (TEAM_KEY / VIEW_KEY เลิกใช้แล้ว ลบออกจาก Script properties ได้)');
  Logger.log(out.join('\n'));
  return out;
}
