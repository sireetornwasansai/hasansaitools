/**
 * HA Round — Apps Script backend
 * ตั้งค่า: Project Settings → Script properties
 *   TEAM_KEY    = รหัสผ่านทีม
 *   GAS_SECRET  = สตริงยาว ๆ สุ่ม (ต้องตรงกับ GAS_SECRET ใน Vercel)
 * ชีตสร้างให้อัตโนมัติ: Issues | Dates | Hait | Files | Log
 *   FOLDER_ID (ไม่บังคับ) = ไอดีโฟลเดอร์ Google Drive ที่จะเก็บไฟล์ ถ้าไม่ใส่จะสร้าง HA-Round-Files ให้
 *   SHARE_MODE (ไม่บังคับ) = private (ค่าเริ่มต้น) | domain | link
 *   MEMO_* (ไม่บังคับ) = ตั้งค่าบันทึกข้อความ ดูหัวข้อ "บันทึกข้อความ" ด้านล่าง
 * ทุก request เป็น POST (text/plain JSON) ผ่าน proxy /api/gas เท่านั้น
 */
var SHEETS = {
  Issues: ['id', 'no', 'deleted', 'updatedAt', 'updatedBy', 'json', 'files'],
  Dates: ['key', 'json'],
  Hait: ['k', 'json'],
  Files: ['id', 'deleted', 'uploadedAt', 'name', 'cat', 'ref', 'note', 'size', 'url', 'driveId', 'by', 'updatedAt', 'mime', 'issue'],
  Log: ['ts', 'user', 'action', 'target', 'diff'],
  Memos: ['id', 'deleted', 'createdAt', 'docNo', 'subject', 'docId', 'pdfId', 'docxId', 'by', 'updatedAt', 'meeting', 'json']
};
var FILE_MAX = 20 * 1024 * 1024;     // ขนาดสูงสุดต่อไฟล์ (ไบต์)
var BAD_EXT = /\.(exe|bat|cmd|com|scr|msi|js|vbs|ps1|sh|jar|apk|dll|lnk)$/i;
var NESTED = ['round', 'backup', 'aft', 'aftStatus'];
function toIso_(v) { return v instanceof Date ? v.toISOString() : String(v || ''); }

var ROLE_ = '';
function roleOf_(key) {
  var P = PropertiesService.getScriptProperties();
  if (key && key === P.getProperty('TEAM_KEY')) return 'editor';
  if (key && P.getProperty('VIEW_KEY') && key === P.getProperty('VIEW_KEY')) return 'viewer';
  return '';
}
function doGet() { return out_({ ok: false, error: 'use_post' }); }

function doPost(e) {
  var b;
  try { b = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'bad_request' }); }
  var P = PropertiesService.getScriptProperties();
  if (!P.getProperty('GAS_SECRET') || b.secret !== P.getProperty('GAS_SECRET')) return out_({ ok: false, error: 'unauthorized' });
  ROLE_ = roleOf_(b.key);
  if (!ROLE_) return out_({ ok: false, error: 'unauthorized' });
  var lock = LockService.getScriptLock();
  var isRead = (b.action === 'ver' || b.action === 'snapshot' || b.action === 'memoGet');
  var extra = null;
  if (!isRead && ROLE_ !== 'editor') return out_({ ok: false, error: 'forbidden' });
  if (!isRead && !lock.tryLock(20000)) return out_({ ok: false, error: 'busy' });
  try {
    var who = String(b.by || '').slice(0, 80);
    switch (b.action) {
      case 'fileChunk': var fr = fileChunk_(b, who); if (!fr.done) return out_({ ok: true, part: b.idx }); break;
      case 'fileUpdate': fileUpdate_(b.id, b.patch || {}, who); break;
      case 'fileDelete': fileDelete_(b.id, who); break;
      case 'ver': return out_({ ok: true, ver: getVer_(), role: ROLE_ });
      case 'snapshot': return out_(snapshot_());
      case 'memoGet': return out_({ ok: true, memo: memoGet_(b.id) });
      case 'memoCreate': extra = memoCreate_(b.memo, who); break;
      case 'memoDelete': memoDelete_(b.id, who); break;
      case 'put': putIssue_(b.issue || {}, who); break;
      case 'patch': patchIssue_(b.id, b.patch || {}, who); break;
      case 'delete': deleteIssue_(b.id, who); break;
      case 'dates': patchDates_(b.patch || {}, who); break;
      case 'hait': setHait_(b.kind, b.key2, b.val, who); break;
      default: return out_({ ok: false, error: 'bad_action' });
    }
    bumpVer_();
    var snap = snapshot_();
    if (extra) snap.result = extra;
    return out_(snap);
  } catch (err) {
    var em = err && err.message; return out_({ ok: false, error: ['notfound', 'badtype', 'toolarge', 'missingchunk', 'folder', 'bad', 'memofail'].indexOf(em) >= 0 ? em : 'error' });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

/* ---------- helpers ---------- */
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function sh_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID'));
  var s = ss.getSheetByName(name);
  if (!s) { s = ss.insertSheet(name); s.appendRow(SHEETS[name]); s.setFrozenRows(1); }
  else if (s.getLastColumn() < SHEETS[name].length && s.getLastRow() >= 1) {   // เพิ่มคอลัมน์ใหม่ในชีตเดิม
    s.getRange(1, 1, 1, SHEETS[name].length).setValues([SHEETS[name]]);
  }
  return s;
}
function rows_(name) { var s = sh_(name), n = s.getLastRow(); return n < 2 ? [] : s.getRange(2, 1, n - 1, SHEETS[name].length).getValues(); }
function getVer_() {
  var c = CacheService.getScriptCache(), v = c.get('VER');
  if (v === null) { v = PropertiesService.getScriptProperties().getProperty('VER') || '0'; c.put('VER', v, 21600); }
  return Number(v);
}
function bumpVer_() {
  var v = String(getVer_() + 1);
  PropertiesService.getScriptProperties().setProperty('VER', v);
  CacheService.getScriptCache().put('VER', v, 21600);
}
function log_(user, action, target, diff) {
  sh_('Log').appendRow([new Date(), user, action, target, String(diff || '').slice(0, 2000)]);
}
function nextNo_() {
  var P = PropertiesService.getScriptProperties(), n = Number(P.getProperty('NEXT_NO') || 0);
  if (!n) { n = rows_('Issues').reduce(function (m, r) { return Math.max(m, Number(r[1]) || 0); }, 0) + 1; }
  P.setProperty('NEXT_NO', String(n + 1));
  return n;
}
function findIssueRow_(id) {
  var s = sh_('Issues'), n = s.getLastRow();
  if (n < 2) return 0;
  var ids = s.getRange(2, 1, n - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (ids[i][0] === id) return i + 2;
  return 0;
}

/* ---------- Issues ---------- */
function putIssue_(issue, who) {
  var s = sh_('Issues'), now = new Date().toISOString();
  if (!issue.id) throw new Error('bad');
  var row = findIssueRow_(issue.id);
  if (row) {                       // แก้ไขของเดิม: คงเลข P ไว้
    var old = JSON.parse(s.getRange(row, 6).getValue() || '{}');
    issue.no = Number(s.getRange(row, 2).getValue()) || old.no;
  } else if (!Number(issue.no)) {  // ใหม่: เซิร์ฟเวอร์ออกเลข
    issue.no = nextNo_();
  }
  var vals = [[issue.id, issue.no, false, now, who, JSON.stringify(issue)]];
  if (row) s.getRange(row, 1, 1, 6).setValues(vals); else s.appendRow(vals[0]);
  log_(who, row ? 'put-update' : 'put-new', issue.id, '');
}
function patchIssue_(id, patch, who) {
  var s = sh_('Issues'), row = findIssueRow_(id);
  if (!row) throw new Error('notfound');
  var cur = JSON.parse(s.getRange(row, 6).getValue() || '{}');
  delete patch.id; delete patch.no;
  Object.keys(patch).forEach(function (k) { cur[k] = patch[k]; });
  var now = new Date().toISOString();
  cur.updatedAt = now; cur.updatedBy = who;
  s.getRange(row, 4, 1, 3).setValues([[now, who, JSON.stringify(cur)]]);
  log_(who, 'patch', id, JSON.stringify(patch));
}
function deleteIssue_(id, who) {      // soft delete
  var s = sh_('Issues'), row = findIssueRow_(id);
  if (!row) throw new Error('notfound');
  s.getRange(row, 3, 1, 3).setValues([[true, new Date().toISOString(), who]]);
  log_(who, 'delete', id, '');
}


/* ---------- Files (เก็บไฟล์ใน Google Drive / รายการใน Sheet 'Files') ---------- */
function cleanName_(s, max) { return String(s || '').replace(/[\\\/\u0000-\u001f<>:"|?*]/g, '_').replace(/^\.+/, '').trim().slice(0, max || 180); }
function rootFolder_() {
  var P = PropertiesService.getScriptProperties(), id = P.getProperty('FOLDER_ID') || P.getProperty('ROOT_FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) { if (P.getProperty('FOLDER_ID')) throw new Error('folder'); } }
  var f = DriveApp.createFolder('HA-Round-Files'); P.setProperty('ROOT_FOLDER_ID', f.getId()); return f;
}
function subFolder_(parent, name) {
  name = cleanName_(name, 60) || 'อื่น ๆ';
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}
function applyShare_(file) {
  var m = PropertiesService.getScriptProperties().getProperty('SHARE_MODE') || 'private';
  try {
    if (m === 'link') file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    else if (m === 'domain') file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {}
}
function findFileRow_(id) {
  var s = sh_('Files'), n = s.getLastRow();
  if (n < 2) return 0;
  var ids = s.getRange(2, 1, n - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (ids[i][0] === id) return i + 2;
  return 0;
}
// รับไฟล์ทีละชิ้น (base64) → ชิ้นสุดท้ายประกอบเป็นไฟล์จริงใน Drive แล้วบันทึกลงชีต Files
function fileChunk_(b, who) {
  var uid = String(b.uid || ''), idx = Number(b.idx), total = Number(b.total);
  if (!/^[A-Za-z0-9_-]{6,40}$/.test(uid) || !(total >= 1 && total <= 12) || !(idx >= 0 && idx < total)) throw new Error('bad');
  var name = cleanName_(b.name);
  if (!name || BAD_EXT.test(name)) throw new Error('badtype');
  var size = Number(b.size) || 0;
  if (size > FILE_MAX) throw new Error('toolarge');
  var data = String(b.data || '');
  if (!data) throw new Error('bad');
  var tmp = subFolder_(rootFolder_(), '_tmp');
  tmp.createFile(uid + '_' + idx, data, MimeType.PLAIN_TEXT);
  if (idx < total - 1) return { done: false };
  // ชิ้นสุดท้าย: ประกอบ
  var all = '';
  for (var i = 0; i < total; i++) {
    var it = tmp.getFilesByName(uid + '_' + i);
    if (!it.hasNext()) throw new Error('missingchunk');
    all += it.next().getBlob().getDataAsString();
  }
  var bytes = Utilities.base64Decode(all);
  if (bytes.length > FILE_MAX) throw new Error('toolarge');
  var blob = Utilities.newBlob(bytes, String(b.mime || 'application/octet-stream').slice(0, 100), name);
  var cat = String(b.cat || 'อื่น ๆ').slice(0, 60);
  var f = subFolder_(rootFolder_(), cat).createFile(blob);
  applyShare_(f);
  for (var j = 0; j < total; j++) { var ti = tmp.getFilesByName(uid + '_' + j); while (ti.hasNext()) ti.next().setTrashed(true); }
  var id = 'f' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6), now = new Date().toISOString();
  var issue = String(b.issue || '').slice(0, 40);
  sh_('Files').appendRow([id, false, now, name, cat, String(b.ref || '').slice(0, 120), String(b.note || '').slice(0, 500), bytes.length, f.getUrl(), f.getId(), who, now, blob.getContentType(), issue]);
  if (issue) syncIssueFiles_(issue);
  log_(who, 'file-upload', id, name + ' (' + bytes.length + ' B)');
  return { done: true, id: id };
}
function fileUpdate_(id, patch, who) {
  var s = sh_('Files'), row = findFileRow_(id);
  if (!row) throw new Error('notfound');
  var v = s.getRange(row, 1, 1, SHEETS.Files.length).getValues()[0], oldIssue = v[13], f = null;
  try { f = DriveApp.getFileById(v[9]); } catch (e) {}
  if (patch.name !== undefined) {
    var nm = cleanName_(patch.name);
    if (!nm || BAD_EXT.test(nm)) throw new Error('badtype');
    if (f) f.setName(nm); v[3] = nm;
  }
  if (patch.cat !== undefined && String(patch.cat) !== String(v[4])) {
    var cat = String(patch.cat).slice(0, 60);
    if (f) f.moveTo(subFolder_(rootFolder_(), cat)); v[4] = cat;
  }
  if (patch.ref !== undefined) v[5] = String(patch.ref).slice(0, 120);
  if (patch.note !== undefined) v[6] = String(patch.note).slice(0, 500);
  if (patch.issue !== undefined) v[13] = String(patch.issue).slice(0, 40);
  v[11] = new Date().toISOString();
  s.getRange(row, 1, 1, SHEETS.Files.length).setValues([v]);
  if (oldIssue) syncIssueFiles_(oldIssue);
  if (v[13] && v[13] !== oldIssue) syncIssueFiles_(v[13]);
  log_(who, 'file-update', id, JSON.stringify(patch));
}
function fileDelete_(id, who) {      // soft delete + ย้ายไฟล์ใน Drive ลงถังขยะ (กู้คืนได้ 30 วัน)
  var s = sh_('Files'), row = findFileRow_(id);
  if (!row) throw new Error('notfound');
  var v = s.getRange(row, 1, 1, SHEETS.Files.length).getValues()[0];
  try { DriveApp.getFileById(v[9]).setTrashed(true); } catch (e) {}
  s.getRange(row, 2).setValue(true);
  s.getRange(row, 12).setValue(new Date().toISOString());
  if (v[13]) syncIssueFiles_(v[13]);
  log_(who, 'file-delete', id, v[3]);
}
// เขียนรายชื่อไฟล์+ลิงก์ลงคอลัมน์ 'files' ของชีต Issues เพื่อให้เปิดดูในชีตได้ตรง ๆ
function syncIssueFiles_(issueId) {
  var row = findIssueRow_(issueId);
  if (!row) return;
  var lines = rows_('Files').filter(function (r) { return r[13] === issueId && r[1] !== true && String(r[1]).toLowerCase() !== 'true'; })
    .map(function (r) { return r[3] + ' : ' + r[8]; });
  sh_('Issues').getRange(row, 7).setValue(lines.join('\n'));
}

/* ---------- Dates / Hait ---------- */
function patchDates_(patch, who) {
  var s = sh_('Dates'), data = rows_('Dates'), idx = {};
  data.forEach(function (r, i) { idx[r[0]] = i + 2; });
  Object.keys(patch).forEach(function (k) {
    var v = patch[k];
    if (NESTED.indexOf(k) >= 0 && v && typeof v === 'object') {
      var cur = idx[k] ? JSON.parse(s.getRange(idx[k], 2).getValue() || '{}') : {};
      Object.keys(v).forEach(function (kk) { if (v[kk] === null || v[kk] === '') delete cur[kk]; else cur[kk] = v[kk]; });
      v = cur;
    }
    if (idx[k]) s.getRange(idx[k], 2).setValue(JSON.stringify(v)); else s.appendRow([k, JSON.stringify(v)]);
  });
  log_(who, 'dates', '', JSON.stringify(patch));
}
function setHait_(kind, key, val, who) {
  if (kind !== 'cat' && kind !== 'plus') throw new Error('bad');
  var s = sh_('Hait'), k = kind + '|' + key, data = rows_('Hait'), row = 0;
  data.forEach(function (r, i) { if (r[0] === k) row = i + 2; });
  if (row) s.getRange(row, 2).setValue(JSON.stringify(val)); else s.appendRow([k, JSON.stringify(val)]);
  log_(who, 'hait', k, JSON.stringify(val));
}

/* ---------- บันทึกข้อความ: สร้าง Google Docs → Word (.docx) + PDF เก็บใน Drive ----------
 * Script properties (ไม่บังคับ):
 *   MEMO_DOCNO_PREFIX = คำนำหน้าเลขที่ เช่น "ชม 0033.301/" (ว่าง = 001/2569)
 *   MEMO_TEMPLATE_ID  = ไอดี Google Docs ต้นแบบ ใช้ {{DOC_NO}} {{DOC_DATE}} {{FROM}} {{TO}} {{SUBJECT}}
 *                       {{ATTACH}} {{CONTENT}} {{NAME}} {{POSITION}} {{CONTACT}} (ไม่ใส่ = ใช้รูปแบบมาตรฐานของระบบ)
 *   MEMO_LOGO_ID      = ไอดีไฟล์ภาพตราครุฑใน Drive (PNG/JPG) วางมุมซ้ายบน
 *   MEMO_FONT         = ชื่อฟอนต์ (ค่าเริ่มต้น Sarabun)
 * หมายเหตุ: doPost ถือ ScriptLock อยู่แล้ว จึงไม่ล็อกซ้ำในนี้ (ล็อกซ้อนจะค้าง)
 */
var MEMO_TYPES = ['invite', 'report', 'follow', 'general'];
var MEMO_CAT = 'คำสั่ง/ประกาศ/หนังสือราชการ';
var THAI_MONTHS_ = ['', 'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];
var DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

function thaiDate_(iso) {
  var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(iso || ''));
  if (!m) throw new Error('bad');
  var y = +m[1], mo = +m[2], d = +m[3];
  if (y < 2000 || y > 2200 || mo < 1 || mo > 12 || d < 1 || d > 31) throw new Error('bad');
  return d + ' ' + THAI_MONTHS_[mo] + ' ' + (y + 543);
}
function memoText_(v, max, multi) {
  var s = String(v == null ? '' : v).replace(/\r\n?/g, '\n').replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f\u200e\u200f\u202a-\u202e]/g, '');
  if (!multi) s = s.replace(/\n+/g, ' ');
  return s.trim().slice(0, max);
}
// ตรวจและทำความสะอาดข้อมูลฝั่งเซิร์ฟเวอร์ทุกครั้ง (ไม่เชื่อค่าจากหน้าเว็บ)
function memoClean_(m) {
  if (!m || typeof m !== 'object') throw new Error('bad');
  var f = {
    type: MEMO_TYPES.indexOf(m.type) >= 0 ? m.type : 'general',
    docNo: memoText_(m.docNo, 40),
    date: memoText_(m.date, 10) || Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd'),
    from: memoText_(m.from, 200),
    to: memoText_(m.to, 200),
    subject: memoText_(m.subject, 300),
    attach: memoText_(m.attach, 300),
    body: memoText_(m.body, 8000, true),
    signName: memoText_(m.signName, 120),
    signPos: memoText_(m.signPos, 160),
    contact: memoText_(m.contact, 200),
    approve: m.approve === true,
    approveTitle: memoText_(m.approveTitle, 160),
    meeting: memoText_(m.meeting, 40),
    reuse: memoText_(m.reuse, 40)
  };
  if (!f.subject || !f.to || !f.body) throw new Error('bad');
  thaiDate_(f.date);
  if (f.body.split('\n').length > 150) throw new Error('bad');
  return f;
}
function memoBlocks_(text) {
  var out = [];
  String(text || '').split('\n').forEach(function (ln) {
    ln = ln.replace(/\s+$/, '');
    if (!ln.trim()) return;
    if (/^\s*(\d+[.)]|\(\d+\)|[-•*])\s+/.test(ln)) out.push({ t: 'li', s: ln.trim().replace(/^[-*]\s+/, '• ') });
    else out.push({ t: 'p', s: ln.trim() });
  });
  return out;
}
// เลขที่หนังสือ: peek ก่อนสร้าง แล้ว commit เมื่อสร้างสำเร็จ (ล้มเหลวไม่เสียเลข) · รีเซ็ตเมื่อขึ้นปี พ.ศ. ใหม่
function memoNoPeek_() {
  var P = PropertiesService.getScriptProperties();
  var year = Number(Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy')) + 543;
  var seq = Number(P.getProperty('MEMO_DOCNO_YEAR')) === year ? Number(P.getProperty('MEMO_DOCNO_SEQ') || 0) : 0;
  seq++;
  var pad = seq < 10 ? '00' + seq : seq < 100 ? '0' + seq : String(seq);
  return { text: (P.getProperty('MEMO_DOCNO_PREFIX') || '') + pad + '/' + year, seq: seq, year: year };
}
function memoNoCommit_(p) {
  var P = PropertiesService.getScriptProperties();
  P.setProperty('MEMO_DOCNO_YEAR', String(p.year));
  P.setProperty('MEMO_DOCNO_SEQ', String(p.seq));
}
function appendRowText_(name, vals, textCols) {
  var s = sh_(name), row = s.getLastRow() + 1;
  (textCols || []).forEach(function (c) { s.getRange(row, c).setNumberFormat('@'); });
  s.getRange(row, 1, 1, vals.length).setValues([vals]);
  return row;
}
function findMemoRow_(id) {
  var s = sh_('Memos'), n = s.getLastRow();
  if (n < 2) return 0;
  var ids = s.getRange(2, 1, n - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (ids[i][0] === id) return i + 2;
  return 0;
}
function registerFile_(file, name, ref, note, who) {
  var id = 'f' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6), now = new Date().toISOString();
  appendRowText_('Files', [id, false, now, name, MEMO_CAT, String(ref || '').slice(0, 120), String(note || '').slice(0, 500), file.getSize(), file.getUrl(), file.getId(), who, now, file.getMimeType(), ''], [6]);
}
function dropFileByDriveId_(driveId, who) {
  if (!driveId) return;
  var hit = rows_('Files').filter(function (r) { return r[9] === driveId && r[1] !== true && String(r[1]).toLowerCase() !== 'true'; })[0];
  if (hit) fileDelete_(hit[0], who);
  else { try { DriveApp.getFileById(driveId).setTrashed(true); } catch (e) {} }
}

/* ----- สร้างเอกสารจากต้นแบบ {{TOKEN}} ----- */
function memoFillTokens_(container, map) {
  Object.keys(map).forEach(function (k) {
    var pat = '\\{\\{' + k + '\\}\\}', val = String(map[k] || '').replace(/\{\{/g, '{ {'), r = container.findText(pat);
    while (r) {
      var el = r.getElement().asText(), st = r.getStartOffset(), en = r.getEndOffsetInclusive();
      el.deleteText(st, en);
      if (val) el.insertText(st, val);
      r = container.findText(pat);
    }
  });
}
function memoFromTemplate_(tplId, f, no, dateText, name, folder) {
  var copy = DriveApp.getFileById(tplId).makeCopy(name, folder), doc = DocumentApp.openById(copy.getId());
  var map = { DOC_NO: no, DOC_DATE: dateText, FROM: f.from, TO: f.to, SUBJECT: f.subject, ATTACH: f.attach, CONTENT: f.body,
    NAME: f.signName, POSITION: f.signPos, CONTACT: f.contact };
  memoFillTokens_(doc.getBody(), map);
  var hd = doc.getHeader(), ft = doc.getFooter();
  if (hd) memoFillTokens_(hd, map);
  if (ft) memoFillTokens_(ft, map);
  doc.saveAndClose();
  return copy.getId();
}

/* ----- สร้างเอกสารรูปแบบบันทึกข้อความมาตรฐาน (ไม่ต้องมีต้นแบบ) ----- */
function memoBuild_(f, no, dateText, name, folder) {
  var P = PropertiesService.getScriptProperties(), font = P.getProperty('MEMO_FONT') || 'Sarabun';
  var doc = DocumentApp.create(name), body = doc.getBody(), id = doc.getId();
  body.setPageWidth(595.28).setPageHeight(841.89).setMarginTop(56.7).setMarginBottom(56.7).setMarginLeft(85).setMarginRight(56.7);
  var base = {};
  base[DocumentApp.Attribute.FONT_FAMILY] = font;
  base[DocumentApp.Attribute.FONT_SIZE] = 16;
  base[DocumentApp.Attribute.FOREGROUND_COLOR] = '#000000';
  base[DocumentApp.Attribute.BOLD] = false;
  var CENTER = DocumentApp.HorizontalAlignment.CENTER, LEFT = DocumentApp.HorizontalAlignment.LEFT;
  function para(text, o) {
    o = o || {};
    var p = body.appendParagraph(text || '');
    p.setAttributes(base);
    p.setLineSpacing(1).setSpacingBefore(o.before || 0).setSpacingAfter(o.after || 0).setAlignment(o.align || LEFT);
    p.setIndentStart(o.start || 0).setIndentFirstLine(o.first || 0);
    if (o.size) p.setFontSize(o.size);
    return p;
  }
  function labelRuns(p, label, value) {
    var a = p.appendText(label + '  '); a.setAttributes(base); a.setBold(true);
    var b = p.appendText(value || ''); b.setAttributes(base); b.setBold(false);
  }
  function labelLine(label, value, o) { var p = para('', o); labelRuns(p, label, value); return p; }
  function noPad(cell) { cell.setPaddingTop(0).setPaddingBottom(0).setPaddingLeft(0).setPaddingRight(4); }
  function cellPara(cell) { var p = cell.getChild(0).asParagraph(); p.setText(''); p.setAttributes(base); return p; }
  function dots(n) { return new Array(n + 1).join('.'); }
  function signBlock(lines, before) {
    var t = body.appendTable([['', '']]); t.setBorderWidth(0);
    var r = t.getRow(0), l = r.getCell(0), c = r.getCell(1);
    noPad(l); noPad(c); l.setWidth(203.6); c.setWidth(250);
    var p0 = cellPara(c); p0.setText(lines[0]); p0.setAttributes(base); p0.setAlignment(CENTER).setSpacingBefore(before || 40).setLineSpacing(1);
    for (var i = 1; i < lines.length; i++) {
      if (!lines[i]) continue;
      var q = c.appendParagraph(lines[i]); q.setAttributes(base); q.setAlignment(CENTER).setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1);
    }
  }

  // หัวกระดาษ: [ตราครุฑ] บันทึกข้อความ [ว่าง]
  var hdr = body.appendTable([['', '', '']]); hdr.setBorderWidth(0);
  var hr = hdr.getRow(0), c0 = hr.getCell(0), c1 = hr.getCell(1), c2 = hr.getCell(2);
  [c0, c1, c2].forEach(noPad);
  c0.setWidth(70); c1.setWidth(313.6); c2.setWidth(70);
  c1.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
  var tp = cellPara(c1); tp.setText('บันทึกข้อความ'); tp.setAttributes(base); tp.setBold(true).setFontSize(29).setAlignment(CENTER);
  cellPara(c0); cellPara(c2);
  var logoId = P.getProperty('MEMO_LOGO_ID');
  if (logoId) {
    try {
      var img = c0.getChild(0).asParagraph().appendInlineImage(DriveApp.getFileById(logoId).getBlob());
      var w = img.getWidth(), hh = img.getHeight();
      if (w > 0 && hh > 0) img.setWidth(56).setHeight(56 * hh / w);
    } catch (e) {}
  }

  labelLine('ส่วนราชการ', f.from, { before: 8 });
  var t2 = body.appendTable([['', '']]); t2.setBorderWidth(0);
  var r2 = t2.getRow(0), a1 = r2.getCell(0), a2 = r2.getCell(1);
  noPad(a1); noPad(a2); a1.setWidth(230); a2.setWidth(223.6);
  labelRuns(cellPara(a1), 'ที่', no); labelRuns(cellPara(a2), 'วันที่', dateText);
  labelLine('เรื่อง', f.subject, { before: 0 });
  body.appendHorizontalRule();
  labelLine('เรียน', f.to, { before: 4 });
  if (f.attach) labelLine('สิ่งที่ส่งมาด้วย', f.attach, { before: 2 });

  memoBlocks_(f.body).forEach(function (b, i) {
    if (b.t === 'li') para(b.s, { before: 2, first: 70.9, start: 92 });
    else para(b.s, { before: i === 0 ? 8 : 6, first: 70.9 });
  });

  signBlock(['(' + (f.signName || dots(36)) + ')', f.signPos || dots(36)], 44);
  if (f.contact) para(f.contact, { before: 14, size: 14 });
  if (f.approve) {
    para('ความเห็น / ข้อสั่งการผู้บังคับบัญชา', { before: 18 }).setBold(true);
    para(dots(88), { before: 6 }); para(dots(88), { before: 2 });
    signBlock(['(' + dots(36) + ')', f.approveTitle || dots(36)], 30);
  }
  // ลบย่อหน้าว่างแรกสุดที่ระบบสร้างให้
  try { var first = body.getChild(0); if (first.getType() === DocumentApp.ElementType.PARAGRAPH && !first.asParagraph().getText() && body.getNumChildren() > 1) body.removeChild(first); } catch (e) {}
  doc.saveAndClose();
  DriveApp.getFileById(id).moveTo(folder);
  return id;
}

function memoCreate_(raw, who) {
  var f = memoClean_(raw), s = sh_('Memos'), N = SHEETS.Memos.length, row = 0, old = null;
  if (f.reuse) {
    row = findMemoRow_(f.reuse);
    if (!row) throw new Error('notfound');
    old = s.getRange(row, 1, 1, N).getValues()[0];
  }
  var no = f.docNo || (old ? String(old[3]) : ''), pending = null;
  if (!no) { pending = memoNoPeek_(); no = pending.text; }
  f.docNo = no;
  var dateText = thaiDate_(f.date), folder = subFolder_(rootFolder_(), MEMO_CAT);
  var stamp = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyyMMdd_HHmmss');
  var base = cleanName_('บันทึกข้อความ_' + f.subject.slice(0, 40) + '_' + stamp, 120);
  var made = [], docId = '', pdf = null, docx = null, docxUrl = '';
  try {
    var tpl = PropertiesService.getScriptProperties().getProperty('MEMO_TEMPLATE_ID');
    if (tpl) { try { docId = memoFromTemplate_(tpl, f, no, dateText, base, folder); } catch (e1) { log_(who, 'memo-template-fail', '', e1 && e1.message); docId = ''; } }
    if (!docId) docId = memoBuild_(f, no, dateText, base, folder);
    made.push(docId);
    var docFile = DriveApp.getFileById(docId);
    applyShare_(docFile);
    pdf = folder.createFile(docFile.getAs('application/pdf')).setName(base + '.pdf'); made.push(pdf.getId());
    applyShare_(pdf);
    try {
      var resp = UrlFetchApp.fetch('https://docs.google.com/document/d/' + docId + '/export?format=docx', { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
      if (resp.getResponseCode() === 200) {
        docx = folder.createFile(resp.getBlob()).setName(base + '.docx'); made.push(docx.getId());
        applyShare_(docx);
      }
    } catch (e2) { log_(who, 'memo-docx-fail', '', e2 && e2.message); }
  } catch (e) {
    log_(who, 'memo-error', '', e && e.message);
    made.forEach(function (i) { try { DriveApp.getFileById(i).setTrashed(true); } catch (x) {} });
    throw new Error(e && e.message === 'folder' ? 'folder' : 'memofail');
  }
  var now = new Date().toISOString(), id = old ? String(old[0]) : 'm' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
  var vals = [id, false, old ? toIso_(old[2]) : now, no, f.subject, docId, pdf.getId(), docx ? docx.getId() : '', who, now, f.meeting, JSON.stringify(f)];
  if (row) { s.getRange(row, 4).setNumberFormat('@'); s.getRange(row, 1, 1, N).setValues([vals]); }
  else appendRowText_('Memos', vals, [4]);
  if (pending) memoNoCommit_(pending);
  registerFile_(pdf, base + '.pdf', no, 'บันทึกข้อความ: ' + f.subject, who);
  if (docx) registerFile_(docx, base + '.docx', no, 'บันทึกข้อความ: ' + f.subject, who);
  if (old) { dropFileByDriveId_(String(old[6]), who); dropFileByDriveId_(String(old[7]), who); try { DriveApp.getFileById(String(old[5])).setTrashed(true); } catch (e3) {} }
  log_(who, old ? 'memo-redo' : 'memo-new', id, no + ' ' + f.subject);
  return { id: id, docNo: no, docId: docId, pdfId: pdf.getId(), docxId: docx ? docx.getId() : '', docUrl: 'https://docs.google.com/document/d/' + docId + '/edit',
    pdfUrl: pdf.getUrl(), docxUrl: docx ? docx.getUrl() : 'https://docs.google.com/document/d/' + docId + '/export?format=docx' };
}
function memoGet_(id) {
  var row = findMemoRow_(String(id || ''));
  if (!row) throw new Error('notfound');
  var v = sh_('Memos').getRange(row, 1, 1, SHEETS.Memos.length).getValues()[0];
  if (v[1] === true || String(v[1]).toLowerCase() === 'true') throw new Error('notfound');
  var form = {}; try { form = JSON.parse(v[11] || '{}'); } catch (e) {}
  return { id: v[0], form: form };
}
function memoDelete_(id, who) {
  var s = sh_('Memos'), row = findMemoRow_(String(id || ''));
  if (!row) throw new Error('notfound');
  var v = s.getRange(row, 1, 1, SHEETS.Memos.length).getValues()[0];
  try { DriveApp.getFileById(String(v[5])).setTrashed(true); } catch (e) {}
  dropFileByDriveId_(String(v[6]), who); dropFileByDriveId_(String(v[7]), who);
  s.getRange(row, 2).setValue(true); s.getRange(row, 10).setValue(new Date().toISOString());
  log_(who, 'memo-delete', id, String(v[3]));
}
function memosSnap_() {
  return rows_('Memos').filter(function (r) { return r[1] !== true && String(r[1]).toLowerCase() !== 'true'; })
    .map(function (r) {
      var j = {}; try { j = JSON.parse(r[11] || '{}'); } catch (e) {}
      return { id: r[0], at: toIso_(r[2]), docNo: String(r[3]), subject: r[4], docId: r[5], pdfId: r[6], docxId: r[7], by: r[8], meeting: r[10] || '', type: j.type || 'general', date: j.date || '', to: j.to || '' };
    }).sort(function (a, b) { return String(b.at).localeCompare(String(a.at)); }).slice(0, 100);
}

/* ---------- snapshot ---------- */
function snapshot_() {
  var issues = rows_('Issues').filter(function (r) { return r[2] !== true && String(r[2]).toLowerCase() !== 'true'; })
    .map(function (r) { try { return JSON.parse(r[5]); } catch (e) { return null; } }).filter(Boolean);
  var dates = {};
  rows_('Dates').forEach(function (r) { try { dates[r[0]] = JSON.parse(r[1]); } catch (e) {} });
  var hait = { cat: {}, plus: {} };
  rows_('Hait').forEach(function (r) {
    var p = String(r[0]).split('|'), kind = p.shift(), key = p.join('|');
    try { hait[kind][key] = JSON.parse(r[1]); } catch (e) {}
  });
  var files = rows_('Files').filter(function (r) { return r[1] !== true && String(r[1]).toLowerCase() !== 'true'; })
    .map(function (r) { return { id: r[0], at: toIso_(r[2]), name: r[3], cat: r[4], ref: r[5], note: r[6], size: Number(r[7]) || 0, url: r[8], by: r[10], mime: r[12], issue: r[13] || '' }; });
  return { ok: true, ver: getVer_(), role: ROLE_, issues: issues, dates: dates, hait: hait, files: files, memos: memosSnap_() };
}

/* รันครั้งเดียวเพื่อตั้ง trigger สำรองข้อมูลรายสัปดาห์ */
function setupWeeklyBackup() {
  ScriptApp.newTrigger('weeklyBackup').timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(2).create();
}
function cleanTmp_() {
  try {
    var it = subFolder_(rootFolder_(), '_tmp').getFiles(), lim = Date.now() - 864e5;
    while (it.hasNext()) { var f = it.next(); if (f.getDateCreated().getTime() < lim) f.setTrashed(true); }
  } catch (e) {}
}
function weeklyBackup() {
  cleanTmp_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.copy('HA-Round backup ' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd'));
}

/* ---------- แจ้งเตือนประเด็นเกินกำหนด (ตั้ง NOTIFY_EMAILS = อีเมลคั่นด้วย ,) ---------- */
function overdueDigest() {
  var to = PropertiesService.getScriptProperties().getProperty('NOTIFY_EMAILS');
  if (!to) return;
  ROLE_ = 'editor';
  var today = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');
  var soon = Utilities.formatDate(new Date(Date.now() + 3 * 864e5), 'Asia/Bangkok', 'yyyy-MM-dd');
  var open = snapshot_().issues.filter(function (i) { return i.status !== 'ปิดแล้ว' && i.due; });
  function line(i) { return 'P' + ('00' + i.no).slice(-3) + ' [' + (i.risk || '-') + '] ' + (i.text || '').slice(0, 80) + ' — ผู้รับผิดชอบ: ' + (i.owner || '-') + ' — ครบกำหนด ' + i.due; }
  var over = open.filter(function (i) { return i.due < today; }), near = open.filter(function (i) { return i.due >= today && i.due <= soon; });
  if (!over.length && !near.length) return;
  MailApp.sendEmail(to, 'HA Round: เกินกำหนด ' + over.length + ' / ใกล้ครบกำหนด ' + near.length,
    (over.length ? 'เกินกำหนด\n' + over.map(line).join('\n') + '\n\n' : '') + (near.length ? 'ครบกำหนดใน 3 วัน\n' + near.map(line).join('\n') : ''));
}
function setupDailyDigest() { ScriptApp.newTrigger('overdueDigest').timeBased().atHour(8).everyDays(1).create(); }

/* ---------- ย้ายข้อมูลจากชีตเก่า: ใส่ชื่อชีตเก่าที่มีแถวหัวคอลัมน์ id,no,by,... แล้วรันครั้งเดียว ---------- */
function migrateFromLegacy(sheetName) {
  ROLE_ = 'editor';
  var src = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName || 'Issues_old');
  if (!src) throw new Error('ไม่พบชีต ' + (sheetName || 'Issues_old'));
  var v = src.getDataRange().getValues(), head = v.shift(), n = 0;
  v.forEach(function (r) {
    var o = {}; head.forEach(function (h, i) { if (h) o[h] = r[i] instanceof Date ? Utilities.formatDate(r[i], 'Asia/Bangkok', 'yyyy-MM-dd') : r[i]; });
    if (!o.id) return;
    if (findIssueRow_(o.id)) return;
    if (!Number(o.no)) o.no = nextNo_();
    sh_('Issues').appendRow([o.id, o.no, false, new Date().toISOString(), 'migrate', JSON.stringify(o)]);
    n++;
  });
  bumpVer_(); log_('migrate', 'migrate', sheetName, n + ' rows');
  return n;
}
