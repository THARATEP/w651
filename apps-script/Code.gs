/**
 * ระบบปรับลดราคาในแบบแสดงรายการ ปริมาณงาน และราคา (BOQ)
 * ตามหนังสือคณะกรรมการวินิจฉัยปัญหาการจัดซื้อจัดจ้างและการบริหารพัสดุภาครัฐ
 * ด่วนที่สุด ที่ กค (กวจ) 0405.3/ว 651 ลงวันที่ 26 กันยายน 2568
 *
 * Backend : Google Apps Script + Google Sheets
 * Create by Mr.Tharatep Thongbao
 */

var APP = {
  NAME: 'ระบบปรับลดราคา BOQ (ว 651)',
  VERSION: '1.1.0',
  // ไอคอนแท็บเบราว์เซอร์ จาก github.com/THARATEP/w651 (ไฟล์ web/icons/icon-64.png)
  ICON_URL: 'https://cdn.jsdelivr.net/gh/THARATEP/w651@main/web/icons/icon-64.png',
  SHEETS: {
    PROJECTS: 'Projects',
    ITEMS: 'Items',
    SETTINGS: 'Settings',
    LOG: 'Log'
  },
  PROJECT_COLS: ['id', 'name', 'workType', 'projectNo', 'agency', 'budget', 'offer', 'contractor',
    'contractNo', 'contractDate', 'awardDate', 'markupPct', 'method', 'sumMode', 'announcedMid', 'pctMode', 'adjRound', 'residualId',
    'sections', 'caseInfo', 'status', 'note', 'midPrice', 'contractPrice', 'pct',
    'createdBy', 'createdAt', 'updatedBy', 'updatedAt'],
  ITEM_COLS: ['projectId', 'seq', 'id', 'section', 'type', 'no', 'name', 'unit', 'qty', 'cost', 'unitCost'],
  SETTING_DEFAULTS: {
    orgName: 'สำนักงานการปฏิรูปที่ดินเพื่อเกษตรกรรมจังหวัดจันทบุรี',
    orgShort: 'ส.ป.ก.จันทบุรี',
    officerName: '',
    officerPosition: '',
    checkerName: '',
    checkerPosition: '',
    headName: '',
    headPosition: 'ปฏิรูปที่ดินจังหวัดจันทบุรี',
    reportFolderId: ''
  }
};

/* ============================ Web entry ============================ */

function doGet(e) {
  ensureSetup_();
  var t = HtmlService.createTemplateFromFile('Index');
  return t.evaluate()
    .setTitle('ปรับลดราคา BOQ ตาม ว 651')
    .setFaviconUrl(APP.ICON_URL.indexOf('<') < 0 ? APP.ICON_URL : 'https://cdn.jsdelivr.net/npm/@tabler/icons@3.19.0/icons/outline/calculator.svg')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

/* ============================ Setup ============================ */

function getDb_() {
  var id = PropertiesService.getScriptProperties().getProperty('DB_ID');
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (err) { /* fall through & recreate */ }
  }
  var active = null;
  try { active = SpreadsheetApp.getActiveSpreadsheet(); } catch (err) { active = null; }
  var ss = active || SpreadsheetApp.create('ฐานข้อมูล ระบบปรับลดราคา BOQ (ว 651)');
  PropertiesService.getScriptProperties().setProperty('DB_ID', ss.getId());
  return ss;
}

/** เรียกครั้งแรกจาก editor เพื่อสร้างชีตฐานข้อมูลและขอสิทธิ์ */
function setup() {
  ensureSetup_(true);
  var ss = getDb_();
  Logger.log('ฐานข้อมูล: ' + ss.getUrl());
  return ss.getUrl();
}

function ensureSetup_(force) {
  var cache = CacheService.getScriptCache();
  if (!force && cache.get('setup_ok')) return;
  var ss = getDb_();
  makeSheet_(ss, APP.SHEETS.PROJECTS, APP.PROJECT_COLS, '#1f5f3f');
  makeSheet_(ss, APP.SHEETS.ITEMS, APP.ITEM_COLS, '#1f5f3f');
  var st = makeSheet_(ss, APP.SHEETS.SETTINGS, ['key', 'value'], '#7a5a12');
  if (st.getLastRow() < 2) {
    var rows = Object.keys(APP.SETTING_DEFAULTS).map(function (k) { return [k, APP.SETTING_DEFAULTS[k]]; });
    st.getRange(2, 1, rows.length, 2).setValues(rows);
  }
  makeSheet_(ss, APP.SHEETS.LOG, ['time', 'user', 'action', 'projectId', 'detail'], '#555555');
  var s1 = ss.getSheetByName('Sheet1') || ss.getSheetByName('แผ่น1');
  if (s1 && ss.getSheets().length > 1 && s1.getLastRow() === 0) ss.deleteSheet(s1);
  cache.put('setup_ok', '1', 21600);
}

function makeSheet_(ss, name, headers, color) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground(color).setFontColor('#ffffff');
    sh.setFrozenRows(1);
  } else {
    // เพิ่มคอลัมน์ใหม่ที่ยังไม่มี (รองรับการอัปเดตเวอร์ชัน)
    var cur = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    headers.forEach(function (h) {
      if (cur.indexOf(h) < 0) {
        sh.getRange(1, sh.getLastColumn() + 1).setValue(h).setFontWeight('bold')
          .setBackground(color).setFontColor('#ffffff');
      }
    });
  }
  return sh;
}

/* ============================ Helpers ============================ */

function userEmail_() {
  try { return Session.getActiveUser().getEmail() || 'ผู้ใช้'; } catch (e) { return 'ผู้ใช้'; }
}

function readTable_(sh) {
  var last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(1, 1, last, sh.getLastColumn()).getValues();
  var head = vals.shift();
  return vals.map(function (r, i) {
    var o = { _row: i + 2 };
    head.forEach(function (h, j) {
      var v = r[j];
      if (v instanceof Date) v = Utilities.formatDate(v, 'Asia/Bangkok', 'yyyy-MM-dd');
      o[h] = v;
    });
    return o;
  });
}

function headerIndex_(sh) {
  var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var m = {};
  head.forEach(function (h, i) { m[h] = i; });
  return { list: head, map: m };
}

function log_(action, projectId, detail) {
  try {
    getDb_().getSheetByName(APP.SHEETS.LOG)
      .appendRow([new Date(), userEmail_(), action, projectId || '', detail || '']);
  } catch (e) { /* ignore */ }
}

function toObjProject_(p) {
  var o = {};
  APP.PROJECT_COLS.forEach(function (k) { o[k] = p[k] === undefined ? '' : p[k]; });
  try { o.sections = p.sections ? JSON.parse(p.sections) : []; } catch (e) { o.sections = []; }
  try { o.caseInfo = p.caseInfo ? JSON.parse(p.caseInfo) : {}; } catch (e) { o.caseInfo = {}; }
  return o;
}

/* ============================ API ============================ */

function apiInit() {
  ensureSetup_();
  return {
    user: userEmail_(),
    settings: apiGetSettings(),
    projects: apiListProjects(),
    dbUrl: getDb_().getUrl(),
    version: APP.VERSION
  };
}

function apiGetSettings() {
  var sh = getDb_().getSheetByName(APP.SHEETS.SETTINGS);
  var o = JSON.parse(JSON.stringify(APP.SETTING_DEFAULTS));
  readTable_(sh).forEach(function (r) { if (r.key) o[r.key] = r.value; });
  return o;
}

function apiSaveSettings(s) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = getDb_().getSheetByName(APP.SHEETS.SETTINGS);
    var rows = readTable_(sh);
    Object.keys(s).forEach(function (k) {
      var hit = rows.filter(function (r) { return r.key === k; })[0];
      if (hit) sh.getRange(hit._row, 2).setValue(s[k]);
      else sh.appendRow([k, s[k]]);
    });
    log_('บันทึกการตั้งค่า', '', '');
    return apiGetSettings();
  } finally { lock.releaseLock(); }
}

function apiListProjects() {
  var sh = getDb_().getSheetByName(APP.SHEETS.PROJECTS);
  return readTable_(sh).map(function (p) {
    return {
      id: p.id, name: p.name, workType: p.workType, projectNo: p.projectNo, agency: p.agency,
      budget: Number(p.budget) || 0, offer: Number(p.offer) || 0, contractor: p.contractor,
      contractNo: p.contractNo, status: p.status, midPrice: Number(p.midPrice) || 0,
      contractPrice: Number(p.contractPrice) || 0, pct: p.pct, updatedAt: p.updatedAt, updatedBy: p.updatedBy
    };
  }).sort(function (a, b) { return String(b.updatedAt).localeCompare(String(a.updatedAt)); });
}

function apiGetProject(id) {
  var db = getDb_();
  var p = readTable_(db.getSheetByName(APP.SHEETS.PROJECTS)).filter(function (r) { return String(r.id) === String(id); })[0];
  if (!p) throw new Error('ไม่พบโครงการ ' + id);
  var items = readTable_(db.getSheetByName(APP.SHEETS.ITEMS))
    .filter(function (r) { return String(r.projectId) === String(id); })
    .sort(function (a, b) { return Number(a.seq) - Number(b.seq); })
    .map(function (r) {
      return { id: String(r.id), section: String(r.section), type: r.type || 'ITEM', no: String(r.no),
        name: r.name, unit: r.unit, qty: r.qty === '' ? '' : Number(r.qty), cost: r.cost === '' ? '' : Number(r.cost),
        unitCost: r.unitCost === '' ? '' : Number(r.unitCost) };
    });
  return { project: toObjProject_(p), items: items };
}

function apiSaveProject(payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var db = getDb_();
    var shP = db.getSheetByName(APP.SHEETS.PROJECTS);
    var shI = db.getSheetByName(APP.SHEETS.ITEMS);
    var p = payload.project;
    var items = payload.items || [];
    var now = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
    var isNew = !p.id;
    if (isNew) {
      p.id = 'P' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyMMddHHmmss') + Math.floor(Math.random() * 90 + 10);
      p.createdBy = userEmail_();
      p.createdAt = now;
    }
    p.updatedBy = userEmail_();
    p.updatedAt = now;

    // คำนวณฝั่งเซิร์ฟเวอร์เพื่อเก็บสรุปผล
    try {
      var res = BoqEngine.compute(p, p.sections || [], items);
      p.midPrice = res.total;
      p.contractPrice = res.sumAdj;
      p.pct = res.pct;
    } catch (e) { /* ignore */ }

    var hi = headerIndex_(shP);
    var row = hi.list.map(function (h) {
      var v = p[h];
      if (h === 'sections' || h === 'caseInfo') return JSON.stringify(v || (h === 'sections' ? [] : {}));
      return v === undefined || v === null ? '' : v;
    });
    var existing = readTable_(shP).filter(function (r) { return String(r.id) === String(p.id); })[0];
    if (existing) {
      // คงค่า createdBy/createdAt เดิม
      row[hi.map.createdBy] = existing.createdBy;
      row[hi.map.createdAt] = existing.createdAt;
      shP.getRange(existing._row, 1, 1, row.length).setValues([row]);
    } else {
      shP.appendRow(row);
    }
    // ลบรายการเดิมของโครงการ แล้วเขียนใหม่
    deleteItems_(shI, p.id);
    if (items.length) {
      var hiI = headerIndex_(shI);
      var rows = items.map(function (it, i) {
        var o = { projectId: p.id, seq: i + 1, id: it.id, section: it.section, type: it.type, no: it.no,
          name: it.name, unit: it.unit, qty: it.qty, cost: it.cost, unitCost: it.unitCost };
        return hiI.list.map(function (h) { return o[h] === undefined || o[h] === null ? '' : o[h]; });
      });
      var start = shI.getLastRow() + 1;
      shI.getRange(start, hiI.map.no + 1, rows.length, 1).setNumberFormat('@');
      shI.getRange(start, 1, rows.length, rows[0].length).setValues(rows);
    }
    log_(isNew ? 'สร้างโครงการ' : 'แก้ไขโครงการ', p.id, p.name);
    return { id: p.id, projects: apiListProjects(), updatedAt: now };
  } finally { lock.releaseLock(); }
}

function deleteItems_(shI, projectId) {
  var last = shI.getLastRow();
  if (last < 2) return;
  var ids = shI.getRange(2, 1, last - 1, 1).getValues();
  // ลบจากล่างขึ้นบน โดยรวมเป็นช่วงต่อเนื่องเพื่อความเร็ว
  var i = ids.length - 1;
  while (i >= 0) {
    if (String(ids[i][0]) === String(projectId)) {
      var end = i;
      while (i - 1 >= 0 && String(ids[i - 1][0]) === String(projectId)) i--;
      shI.deleteRows(i + 2, end - i + 1);
    }
    i--;
  }
}

function apiDeleteProject(id) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var db = getDb_();
    var shP = db.getSheetByName(APP.SHEETS.PROJECTS);
    var hit = readTable_(shP).filter(function (r) { return String(r.id) === String(id); })[0];
    if (hit) shP.deleteRow(hit._row);
    deleteItems_(db.getSheetByName(APP.SHEETS.ITEMS), id);
    log_('ลบโครงการ', id, hit ? hit.name : '');
    return apiListProjects();
  } finally { lock.releaseLock(); }
}

function apiDuplicateProject(id) {
  var d = apiGetProject(id);
  d.project.id = '';
  d.project.name = d.project.name + ' (สำเนา)';
  d.project.contractNo = '';
  return apiSaveProject(d);
}

/* ============================ Export ============================ */

function reportFolder_() {
  var s = apiGetSettings();
  if (s.reportFolderId) {
    try { return DriveApp.getFolderById(s.reportFolderId); } catch (e) { /* create new */ }
  }
  var f = DriveApp.createFolder('รายงานปรับลดราคา BOQ (ว 651)');
  apiSaveSettings({ reportFolderId: f.getId() });
  return f;
}

/**
 * ส่งออกรายงานเป็น Google Sheets (2 แผ่นงาน) และเลือกสร้าง PDF เก็บใน Drive
 * @param {string} id รหัสโครงการ
 * @param {boolean} withPdf สร้าง PDF ด้วยหรือไม่
 */
function apiExportReport(id, withPdf) {
  var d = apiGetProject(id);
  var p = d.project;
  var res = BoqEngine.compute(p, p.sections, d.items);
  var st = apiGetSettings();
  var title = 'รายงานปรับลดราคา - ' + (p.name || p.id);
  var ss = SpreadsheetApp.create(title);
  var sh1 = ss.getSheets()[0].setName('สรุปราคากลาง');
  var sh2 = ss.insertSheet('แนบท้ายสัญญา');
  writeSummarySheet_(sh1, p, res, st);
  writeContractSheet_(sh2, p, res, st);
  SpreadsheetApp.flush();
  var file = DriveApp.getFileById(ss.getId());
  var folder = reportFolder_();
  file.moveTo(folder);
  var out = { sheetUrl: ss.getUrl(), folderUrl: folder.getUrl() };
  if (withPdf) {
    var pdf1 = exportSheetPdf_(ss.getId(), sh1.getSheetId(), false, p.name + ' - แบบสรุปราคากลาง.pdf');
    var pdf2 = exportSheetPdf_(ss.getId(), sh2.getSheetId(), true, p.name + ' - แบบแสดงรายการฯ แนบท้ายสัญญา.pdf');
    folder.createFile(pdf1);
    var f2 = folder.createFile(pdf2);
    out.pdfUrl = f2.getUrl();
  }
  log_('ส่งออกรายงาน', id, withPdf ? 'Sheets+PDF' : 'Sheets');
  return out;
}

function exportSheetPdf_(ssId, gid, portrait, name) {
  var url = 'https://docs.google.com/spreadsheets/d/' + ssId + '/export?format=pdf' +
    '&size=A4&portrait=' + portrait + '&fitw=true&gridlines=false&printtitle=false' +
    '&sheetnames=false&pagenum=CENTER&fzr=true&top_margin=0.5&bottom_margin=0.5' +
    '&left_margin=0.4&right_margin=0.4&gid=' + gid;
  var resp = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) throw new Error('สร้าง PDF ไม่สำเร็จ (' + resp.getResponseCode() + ')');
  return resp.getBlob().setName(name);
}

var NUM = '#,##0.00';

function setTitleRows_(sh, lines, width) {
  lines.forEach(function (t, i) {
    var r = sh.getRange(i + 1, 1, 1, width).merge().setValue(t).setHorizontalAlignment('center');
    if (i === 0) r.setFontWeight('bold').setFontSize(13);
  });
  return lines.length + 1;
}

function styleTable_(sh, r1, c1, nr, nc) {
  sh.getRange(r1, c1, nr, nc).setBorder(true, true, true, true, true, true, '#444444', SpreadsheetApp.BorderStyle.SOLID)
    .setVerticalAlignment('middle');
}

function writeSummarySheet_(sh, p, res, st) {
  var mk = res.markupPct;
  var W = mk ? 10 : 9;
  sh.getRange('A:Z').setFontFamily('Sarabun').setFontSize(10);
  var lines = ['แบบสรุปราคากลางงานก่อสร้าง' + workTypeText_(p)];
  if (mk) lines.push('(กรณีงานก่อสร้างใน 3 จังหวัดชายแดนภาคใต้ ได้แก่ จังหวัดยะลา จังหวัดปัตตานี และจังหวัดนราธิวาส)');
  lines.push('โครงการ ' + (p.name || '') + '  งบประมาณ ' + fmt_(p.budget) + ' บาท  เลขที่โครงการ ' + (p.projectNo || '-'));
  var r = setTitleRows_(sh, lines, W);

  // ตารางสรุปราคากลาง
  sh.getRange(r, 1).setValue('ตาราง : สรุปราคากลางโครงการงานก่อสร้าง').setFontWeight('bold');
  r++;
  var head = ['ลำดับที่', 'รายการ', 'หน่วย', 'ค่างานต้นทุน', 'Factor F / Vat', 'ราคากลาง'];
  if (mk) head.push('ราคากลางบวกเพิ่ม ร้อยละ ' + mk);
  var sumRows = [head];
  p.sections.forEach(function (s, i) {
    var t = res.secTotals[s.code] || {};
    var row = [i + 1, s.name, 'บาท', t.cost || 0, Number(s.factor) || 1, t.mid || 0];
    if (mk) row.push(t.midPlus || 0);
    sumRows.push(row);
  });
  var tot = ['', '', '', '', 'รวม', mk ? '' : res.total];
  if (mk) tot.push(res.total);
  sumRows.push(tot);
  sh.getRange(r, 1, sumRows.length, head.length).setValues(sumRows);
  sh.getRange(r, 1, 1, head.length).setFontWeight('bold').setBackground('#e8f1ec').setHorizontalAlignment('center');
  sh.getRange(r + 1, 4, sumRows.length - 1, 1).setNumberFormat(NUM);
  sh.getRange(r + 1, 5, sumRows.length - 1, 1).setNumberFormat('0.0000');
  sh.getRange(r + 1, 6, sumRows.length - 1, head.length - 5).setNumberFormat(NUM);
  styleTable_(sh, r, 1, sumRows.length - 1, head.length);
  r += sumRows.length + 1;

  sh.getRange(r, 1).setValue('ตาราง : ราคาตามคำเสนอของผู้ชนะการจัดซื้อจัดจ้าง').setFontWeight('bold');
  r++;
  sh.getRange(r, 1, 2, 4).setValues([['ลำดับที่', 'รายการ', 'หน่วย', 'ราคาที่เสนอ'], [1, 'ราคาตามคำเสนอของผู้ชนะการจัดซื้อจัดจ้าง', 'บาท', res.offer]]);
  sh.getRange(r, 1, 1, 4).setFontWeight('bold').setBackground('#e8f1ec');
  sh.getRange(r + 1, 4).setNumberFormat(NUM).setBackground('#dddddd');
  styleTable_(sh, r, 1, 2, 4);
  r += 3;

  var calc = [
    ['วิธีการคำนวณราคากลางหักค่าใช้จ่ายพิเศษตามข้อกำหนด', '', ''],
    ['ราคากลาง', res.total, ''],
    ['หัก ค่าใช้จ่ายพิเศษตามข้อกำหนด', res.special, ''],
    ['ผลต่างระหว่างราคากลางหักค่าใช้จ่ายพิเศษตามข้อกำหนด', res.P1, ''],
    ['', '', ''],
    ['วิธีการคำนวณราคาตามคำเสนอฯ หักค่าใช้จ่ายพิเศษตามข้อกำหนด', '', ''],
    ['ราคาที่เสนอของผู้ชนะการจัดซื้อจัดจ้าง', res.offer, ''],
    ['หัก ค่าใช้จ่ายพิเศษตามข้อกำหนด', res.special, ''],
    ['ผลต่างระหว่างราคาตามคำเสนอฯ หักค่าใช้จ่ายพิเศษตามข้อกำหนด', res.P2, ''],
    ['', '', ''],
    ['ผลต่างระหว่างราคากลางกับราคาตามคำเสนอฯ หักค่าใช้จ่ายพิเศษตามข้อกำหนด', '', ''],
    ['ผลต่างระหว่างราคากลางหักค่าใช้จ่ายพิเศษตามข้อกำหนด', res.P1, '(1)'],
    ['หัก ผลต่างระหว่างราคาที่เสนอฯ หักค่าใช้จ่ายพิเศษตามข้อกำหนด', res.P2, '(2)'],
    ['ผลต่างระหว่างราคากลางกับราคาตามคำเสนอฯ หักค่าใช้จ่ายพิเศษตามข้อกำหนด', res.P3, '(3)'],
    ['ผลต่างตาม (3) เทียบกับราคากลางหักค่าใช้จ่ายพิเศษตามข้อกำหนด (1) คิดเป็นอัตราร้อยละ', res.pct, '']
  ];
  calc.forEach(function (c, i) {
    sh.getRange(r + i, 2).setValue(c[0]);
    if (c[1] !== '') sh.getRange(r + i, 6).setValue(c[1]).setNumberFormat(NUM);
    if (c[2]) sh.getRange(r + i, 7).setValue(c[2]);
  });
  [0, 5, 10].forEach(function (k) { sh.getRange(r + k, 2).setFontWeight('bold').setFontLine('underline'); });
  [3, 8, 13].forEach(function (k) { sh.getRange(r + k, 6).setBackground('#dddddd').setFontWeight('bold'); });
  sh.getRange(r + 14, 6).setBackground('#fff2cc').setFontWeight('bold').setFontSize(12);
  r += calc.length + 1;

  // ตารางรายละเอียด
  var dh1 = ['ลำดับที่', 'รายการ', 'หน่วย', 'จำนวน', 'ราคากลางงานก่อสร้าง (บาท)', '', ''];
  var dh2 = ['', '', '', '', 'ค่างานต้นทุน', 'Factor F / Vat', 'ราคากลาง'];
  if (mk) { dh1.push(''); dh2.push('ราคากลางบวกเพิ่ม ร้อยละ ' + mk); }
  dh1.push('ราคาที่ปรับลดตามคำเสนอ (บาท)', ''); dh2.push('ราคาที่ปรับลด', 'ราคาต่อหน่วย');
  var body = [dh1, dh2];
  var fmtRows = [];
  var rowsOut = detailRows_(p, res, true);
  rowsOut.forEach(function (x) {
    var row = [x.no, x.name, x.unit, x.qty, x.cost, x.factor, x.mid];
    if (mk) row.push(x.midPlus);
    row.push(x.adj, x.unitAdj);
    body.push(row);
    fmtRows.push(x.kind);
  });
  var fin = ['', '', '', '', '', 'วงเงินรวม', res.totalMid];
  if (mk) fin.push(res.total);
  fin.push(res.sumAdj, '(*)');
  body.push(fin);
  var nc = body[0].length;
  sh.getRange(r, 1, body.length, nc).setValues(body);
  sh.getRange(r, 1, 2, nc).setFontWeight('bold').setBackground('#e8f1ec').setHorizontalAlignment('center').setWrap(true);
  sh.getRange(r, 5, 1, mk ? 4 : 3).merge();
  sh.getRange(r, mk ? 9 : 8, 1, 2).merge();
  sh.getRange(r + 2, 4, body.length - 2, 1).setNumberFormat(NUM);
  sh.getRange(r + 2, 5, body.length - 2, 1).setNumberFormat(NUM);
  sh.getRange(r + 2, 6, body.length - 2, 1).setNumberFormat('0.0000');
  sh.getRange(r + 2, 7, body.length - 2, nc - 6).setNumberFormat(NUM);
  fmtRows.forEach(function (k, i) {
    var rg = sh.getRange(r + 2 + i, 1, 1, nc);
    if (k === 'H1' || k === 'SEC') rg.setFontWeight('bold');
    if (k === 'SUB' || k === 'SECTOT') rg.setFontWeight('bold').setBackground('#f3f3f3');
  });
  sh.getRange(r + body.length - 1, 1, 1, nc).setFontWeight('bold');
  sh.getRange(r + body.length - 1, nc - 1).setBackground('#dddddd');
  styleTable_(sh, r, 1, body.length - 1, nc);
  r += body.length + 1;
  sh.getRange(r, 1).setValue('หมายเหตุ');
  sh.getRange(r, 2).setValue('(*) ส่วนต่างนี้ให้ปรับเพิ่มลด ในรายการใดรายการหนึ่งได้ตามความเหมาะสม ยกเว้น รายการค่าใช้จ่ายพิเศษตามข้อกำหนด');
  sh.getRange(r + 1, 2).setValue('ทั้งนี้ วงเงินรวมของการปรับลดราคาจะต้องเท่ากับราคาตามคำเสนอของผู้ชนะการจัดซื้อจัดจ้าง');
  r += 3;
  signBlock_(sh, r, [[st.officerName, st.officerPosition, 'ผู้ปรับลดราคา (เจ้าหน้าที่ผู้ได้รับมอบหมาย)'],
    [st.checkerName, st.checkerPosition, 'ผู้ตรวจสอบ']], nc);
  colWidths_(sh, mk ? [55, 330, 55, 75, 105, 75, 110, 115, 110, 95] : [55, 330, 55, 75, 105, 75, 110, 110, 95]);
}

function writeContractSheet_(sh, p, res, st) {
  sh.getRange('A:Z').setFontFamily('Sarabun').setFontSize(10);
  var W = 6;
  var tl = ['แบบแสดงรายการ ปริมาณงาน และราคา', 'เอกสารแนบท้ายสัญญาจ้างก่อสร้าง : งานก่อสร้าง' + workTypeText_(p)];
  if (res.markupPct) tl.push('(กรณีงานก่อสร้างใน 3 จังหวัดชายแดนภาคใต้ ได้แก่ จังหวัดยะลา จังหวัดปัตตานี และจังหวัดนราธิวาส)');
  tl.push('เลขที่ ' + (p.contractNo || '..........') + '  ลงวันที่ ' + thaiDate_(p.contractDate),
    'โครงการ ' + (p.name || '') + '  เลขที่โครงการ ' + (p.projectNo || '-'));
  var r = setTitleRows_(sh, tl, W);
  var info = [['หน่วยงานเจ้าของโครงการ', p.agency || st.orgName],
    ['วงเงินงบประมาณ', fmt_(p.budget) + ' บาท'],
    ['ราคากลางงานก่อสร้าง', fmt_(res.total) + ' บาท'],
    ['วงเงินตามสัญญา', fmt_(res.sumAdj) + ' บาท (' + bahtText_(res.sumAdj) + ')'],
    ['ชื่อผู้รับจ้าง', p.contractor || '']];
  info.forEach(function (x, i) {
    sh.getRange(r + i, 1, 1, 2).merge().setValue(x[0]).setFontWeight('bold');
    sh.getRange(r + i, 3, 1, 4).merge().setValue(x[1]);
  });
  r += info.length + 1;
  var body = [['ลำดับที่', 'รายการ', 'หน่วย', 'จำนวน', 'ราคาที่ปรับลดตามคำเสนอ (บาท)', ''],
    ['', '', '', '', 'ราคา', 'ราคาต่อหน่วย']];
  var kinds = [];
  detailRows_(p, res, false).forEach(function (x) {
    body.push([x.no, x.name, x.unit, x.qty, x.adj, x.unitAdj]);
    kinds.push(x.kind);
  });
  body.push(['', '', '', 'วงเงินตามสัญญา', res.sumAdj, '']);
  sh.getRange(r, 1, body.length, W).setValues(body);
  sh.getRange(r, 1, 2, W).setFontWeight('bold').setBackground('#e8f1ec').setHorizontalAlignment('center');
  sh.getRange(r, 5, 1, 2).merge();
  sh.getRange(r + 2, 4, body.length - 2, 3).setNumberFormat(NUM);
  kinds.forEach(function (k, i) {
    var rg = sh.getRange(r + 2 + i, 1, 1, W);
    if (k === 'H1' || k === 'SEC') rg.setFontWeight('bold');
    if (k === 'SUB' || k === 'SECTOT') rg.setFontWeight('bold').setBackground('#f3f3f3');
  });
  sh.getRange(r + body.length - 1, 1, 1, W).setFontWeight('bold');
  sh.getRange(r + body.length - 1, 5).setBackground('#dddddd');
  styleTable_(sh, r, 1, body.length - 1, W);
  r += body.length + 2;
  signBlock_(sh, r, [[st.headName, st.headPosition, 'ผู้ว่าจ้าง'], [p.contractor, '', 'ผู้รับจ้าง']], W);
  colWidths_(sh, [55, 380, 60, 80, 120, 110]);
}

function workTypeText_(p) { return p.workType && p.workType !== 'อื่น ๆ' ? p.workType : ''; }

function signBlock_(sh, r, people, nc) {
  var half = Math.max(1, Math.floor(nc / 2));
  people.forEach(function (pp, i) {
    var c = i === 0 ? 1 : half + 1;
    var w = i === 0 ? half : nc - half;
    sh.getRange(r, c, 1, w).merge().setValue('ลงชื่อ ........................................ ' + pp[2]).setHorizontalAlignment('center');
    sh.getRange(r + 1, c, 1, w).merge().setValue('( ' + (pp[0] || '........................................') + ' )').setHorizontalAlignment('center');
    sh.getRange(r + 2, c, 1, w).merge().setValue(pp[1] ? 'ตำแหน่ง ' + pp[1] : 'ตำแหน่ง ........................................').setHorizontalAlignment('center');
  });
}

function colWidths_(sh, ws) { ws.forEach(function (w, i) { sh.setColumnWidth(i + 1, w); }); }

/** แปลงผลคำนวณเป็นแถวรายงาน (รวมบรรทัดรวมย่อย/รวมหมวด) */
function detailRows_(p, res, full) {
  var out = [];
  var secByCode = {};
  p.sections.forEach(function (s) { secByCode[s.code] = s; });
  var bySec = {};
  res.rows.forEach(function (r) { (bySec[r.section] = bySec[r.section] || []).push(r); });
  p.sections.forEach(function (s, si) {
    var rows = bySec[s.code] || [];
    if (!rows.length) return;
    out.push({ kind: 'SEC', no: String(si + 1), name: s.name, unit: '', qty: '', cost: '', factor: '', mid: '', midPlus: '', adj: '', unitAdj: '' });
    var openH1 = null;
    var flush = function () {
      if (openH1 && openH1.sub && openH1.showSub !== false) {
        out.push({ kind: 'SUB', no: '', name: 'รวมค่า' + String(openH1.name).replace(/^ค่า/, ''), unit: '', qty: '', cost: openH1.sub.cost,
          factor: Number(s.factor) || '', mid: openH1.sub.mid, midPlus: openH1.sub.midPlus, adj: openH1.sub.adj, unitAdj: '' });
      }
      openH1 = null;
    };
    rows.forEach(function (r) {
      if (r.type === 'H1') { flush(); openH1 = r; out.push({ kind: 'H1', no: r.no || '', name: r.name, unit: '', qty: '', cost: '', factor: '', mid: '', midPlus: '', adj: '', unitAdj: '' }); return; }
      if (r.type === 'H2') { out.push({ kind: 'H2', no: r.no || '', name: r.name, unit: '', qty: '', cost: '', factor: '', mid: '', midPlus: '', adj: '', unitAdj: '' }); return; }
      out.push({ kind: 'ITEM', no: r.no || '', name: r.name, unit: r.unit || '', qty: r.qty, cost: r.cost, factor: r.factor,
        mid: r.mid, midPlus: r.midPlus, adj: r.adj, unitAdj: r.unitAdj });
    });
    flush();
    var t = res.secTotals[s.code];
    out.push({ kind: 'SECTOT', no: '', name: 'รวม ' + s.name, unit: '', qty: '', cost: t.cost, factor: Number(s.factor) || '',
      mid: t.mid, midPlus: t.midPlus, adj: t.adj, unitAdj: '' });
  });
  return out;
}

function fmt_(n) { return BoqEngine.fmt(Number(n) || 0); }
function thaiDate_(s) { return BoqEngine.thaiDate(s) || '..........'; }
function bahtText_(n) { return BoqEngine.bahtText(n); }
