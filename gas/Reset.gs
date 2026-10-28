/**
 * HA Round — รีเซตข้อมูลเก่า (รันด้วยมือจากหน้าแก้ไข Apps Script เท่านั้น ไม่ผูกกับเว็บ)
 * 1) (ไม่บังคับ) trashDriveFilesFromSheet()  ย้ายไฟล์แนบ/เอกสารที่ระบบเคยสร้างไปถังขยะ Drive (กู้ได้ 30 วัน)
 * 2) resetAllData()  สำรองชีตเป็นสำเนาใน Drive ก่อน แล้วล้างข้อมูลทุกแท็บ (เหลือแถวหัวตาราง) และรีเซตเลขที่
 * แผนงาน (public/data/plan.json) ไม่ถูกแตะ — เป็นไฟล์บน GitHub
 */
/** ตรวจว่าระบบต่อกับชีตและโฟลเดอร์ไหน — รันแล้วดูที่ Execution log */
function checkConnection() {
  var ss = openSS_(), P = PropertiesService.getScriptProperties();
  Logger.log('ชีตที่ใช้: ' + ss.getName() + ' → ' + ss.getUrl());
  Logger.log('SHEET_ID ตั้งไว้: ' + (P.getProperty('SHEET_ID') ? 'ใช่' : 'ไม่ (กำลังใช้ชีตที่ผูกกับสคริปต์)'));
  var fid = P.getProperty('FOLDER_ID') || P.getProperty('ROOT_FOLDER_ID');
  Logger.log('โฟลเดอร์ไฟล์แนบ: ' + (fid ? DriveApp.getFolderById(fid).getUrl() : 'ยังไม่มี (จะสร้าง HA-Round-Files ตอนอัปโหลดครั้งแรก)'));
}

function backupBeforeReset() {
  var ss = openSS_();
  var name = ss.getName() + ' — สำรองก่อนรีเซต ' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HHmm');
  var copy = DriveApp.getFileById(ss.getId()).makeCopy(name);
  Logger.log('สำรองแล้ว: ' + copy.getUrl());
  return copy.getUrl();
}

function resetAllData() {
  var url = backupBeforeReset();                      // ถ้าสำรองไม่สำเร็จ จะหยุดตรงนี้ ไม่ล้างข้อมูล
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    Object.keys(SHEETS).forEach(function (name) {     // SHEETS มาจาก Code.gs: Issues, Dates, Hait, Files, Log, Memos
      var s = sh_(name), last = s.getLastRow();
      if (last >= 2) s.getRange(2, 1, last - 1, Math.max(s.getLastColumn(), SHEETS[name].length)).clearContent();
    });
    var P = PropertiesService.getScriptProperties();
    ['NEXT_NO', 'MEMO_DOCNO_SEQ', 'MEMO_DOCNO_YEAR'].forEach(function (k) { P.deleteProperty(k); });   // เลขประเด็น P และเลขที่บันทึกข้อความเริ่มใหม่
    bumpVer_();                                       // ให้หน้าเว็บทุกเครื่องดึงข้อมูลใหม่
    Logger.log('ล้างข้อมูลเรียบร้อย · สำรองไว้ที่ ' + url);
  } finally { lock.releaseLock(); }
}

function trashDriveFilesFromSheet() {
  var n = 0;
  function trash(id) { if (!id) return; try { DriveApp.getFileById(String(id)).setTrashed(true); n++; } catch (e) { } }
  rows_('Files').forEach(function (r) { trash(r[9]); });                                 // driveId
  rows_('Memos').forEach(function (r) { trash(r[5]); trash(r[6]); trash(r[7]); });       // docId, pdfId, docxId
  Logger.log('ย้ายไปถังขยะ ' + n + ' ไฟล์');
}