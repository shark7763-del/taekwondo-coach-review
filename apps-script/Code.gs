/**********************************************************************
 * 教練課後復盤系統 — GAS 後端（Google Sheets 當資料庫）
 * 模式：Web App 服務 HTML + google.script.run，前端離線優先、背景同步。
 *
 * 部署前只要做一件事：把下面 SS_ID 換成你的試算表 ID（網址中 /d/ 與 /edit 之間那串）。
 * 留空字串則自動用「綁定的試算表」（把這份腳本從試算表的 擴充功能>Apps Script 開啟時）。
 **********************************************************************/

var SS_ID = '';   // ← 貼上你的 Google 試算表 ID，或留空用綁定試算表

var SHEETS = {
  reviews: 'Reviews',
  coaches: 'Coaches',
  classes: 'Classes',
  contents: 'Contents',
  schedule: 'Schedule',
  students: 'Students'
};
var SCHEDULE_HEADERS = ['id', 'weekday', 'classType', 'start', 'end', 'coach', 'content'];
var STUDENT_HEADERS = ['id', 'name', 'classType', 'active'];

// Reviews 分頁的欄位：前面是可在 Sheet 直接看/篩的扁平欄，最後一欄 json 存完整物件
var REVIEW_HEADERS = [
  'id','date','weekday','coach','classType','studentCount',
  'overall','grade','teaching','control','interaction','attitude',
  'hasFlag','needHelp','safety','problem','improvement','createdAt','json'
];

/* ---------- Web App 入口 ---------- */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('教練課後復盤系統')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ---------- 試算表工具 ---------- */
function ss_() {
  return SS_ID ? SpreadsheetApp.openById(SS_ID) : SpreadsheetApp.getActiveSpreadsheet();
}
function sheet_(name, headers) {
  var ss = ss_();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (headers) sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  } else if (headers) {
    // 確保表頭存在且包含所有欄位（例如後來新增的 content 欄會自動補上）
    var cur = sh.getRange(1, 1, 1, headers.length).getValues()[0];
    var diff = false;
    for (var i = 0; i < headers.length; i++) { if (String(cur[i]) !== headers[i]) { diff = true; break; } }
    if (diff) sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  }
  return sh;
}
function listCol_(name, fallback) {
  var sh = sheet_(name, ['name']);
  var last = sh.getLastRow();
  if (last < 2) {
    if (fallback && fallback.length) {
      sh.getRange(2, 1, fallback.length, 1).setValues(fallback.map(function (x) { return [x]; }));
      return fallback;
    }
    return [];
  }
  return sh.getRange(2, 1, last - 1, 1).getValues().map(function (r) { return r[0]; }).filter(String);
}

/* ---------- 前端啟動時抓全部資料 ---------- */
function getBootData() {
  var coaches = listCol_(SHEETS.coaches, ['楊復傑', '林伊辰', '陸呂香', '劉品妡', '廖彥博', '李仲寧']);
  var classes = listCol_(SHEETS.classes, ['選手班', '健身班', '週六班', '幼幼班', '大同國小社團', '彭福國小社團', '育林國小社團']);
  var contents = listCol_(SHEETS.contents, ['暖身/體能', '基本動作', '前踢', '旋踢', '側踢', '下壓踢', '後踢', '腳靶練習', '品勢', '對打/實戰', '護具對練', '柔軟度', '核心訓練', '考帶複習', '比賽培訓', '禮儀/精神']);
  var schedule = getSchedule_();
  var students = getStudents_();
  var reviews = getReviews_(BOOT_REVIEW_LIMIT);
  return JSON.stringify({ coaches: coaches, classes: classes, contents: contents, schedule: schedule, students: students, reviews: reviews });
}

var BOOT_REVIEW_LIMIT = 1500;  // 開機載入的最大復盤筆數（只取最新的；歷史仍完整保留在 Sheet）
function getReviews_(maxRows) {
  var sh = sheet_(SHEETS.reviews, REVIEW_HEADERS);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var jsonCol = REVIEW_HEADERS.indexOf('json') + 1;
  var total = last - 1;
  var n = (maxRows && maxRows < total) ? maxRows : total;
  var startRow = last - n + 1;   // 只讀最後 n 列（最新的紀錄）
  var vals = sh.getRange(startRow, jsonCol, n, 1).getValues();
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    if (vals[i][0]) { try { out.push(JSON.parse(vals[i][0])); } catch (e) {} }
  }
  return out;
}

/* ---------- 新增 / 覆寫一筆復盤 ---------- */
function saveReviewSrv(reviewJson) {
  var r = (typeof reviewJson === 'string') ? JSON.parse(reviewJson) : reviewJson;
  var sh = sheet_(SHEETS.reviews, REVIEW_HEADERS);
  var cat = r.cat || {};
  var row = REVIEW_HEADERS.map(function (h) {
    if (h === 'json') return JSON.stringify(r);
    if (h === 'hasFlag') return (r.flags || []).some(function (f) { return f !== 'none'; }) ? 'Y' : '';
    if (h === 'needHelp') return r.help ? 'Y' : '';
    if (h === 'safety') return r.safety ? 'Y' : '';
    if (h === 'teaching' || h === 'control' || h === 'interaction' || h === 'attitude') return cat[h] != null ? cat[h] : '';
    return r[h] != null ? r[h] : '';
  });
  // 若 id 已存在則覆寫，否則新增
  var idCol = REVIEW_HEADERS.indexOf('id') + 1;
  var last = sh.getLastRow();
  var found = 0;
  if (last >= 2) {
    var ids = sh.getRange(2, idCol, last - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) { if (ids[i][0] === r.id) { found = i + 2; break; } }
  }
  if (found) {
    sh.getRange(found, 1, 1, REVIEW_HEADERS.length).setValues([row]);   // 覆寫不重推
  } else {
    sh.appendRow(row);
    if (r.lineReport) { try { pushToLine_(r.lineReport); } catch (e) {} } // 新的一筆才自動推 LINE
  }
  return 'ok';
}

/* ---------- 刪除一筆 ---------- */
function deleteReviewSrv(id) {
  var sh = sheet_(SHEETS.reviews, REVIEW_HEADERS);
  var idCol = REVIEW_HEADERS.indexOf('id') + 1;
  var last = sh.getLastRow();
  if (last < 2) return 'ok';
  var ids = sh.getRange(2, idCol, last - 1, 1).getValues();
  for (var i = ids.length - 1; i >= 0; i--) { if (ids[i][0] === id) sh.deleteRow(i + 2); }
  return 'ok';
}

/* ---------- 覆寫教練 / 班別名單 ---------- */
function saveCoachesSrv(arr) { return writeList_(SHEETS.coaches, arr); }
function saveClassesSrv(arr) { return writeList_(SHEETS.classes, arr); }
function saveContentsSrv(arr) { return writeList_(SHEETS.contents, arr); }

/* ---------- 課表讀寫 ---------- */
function getSchedule_() {
  var sh = sheet_(SHEETS.schedule, SCHEDULE_HEADERS);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(2, 1, last - 1, SCHEDULE_HEADERS.length).getValues();
  return vals.filter(function (r) { return r[0]; }).map(function (r) {
    var o = {}; SCHEDULE_HEADERS.forEach(function (h, i) { o[h] = r[i] != null ? String(r[i]) : ''; }); return o;
  });
}
function saveScheduleSrv(arr) {
  var list = (typeof arr === 'string') ? JSON.parse(arr) : arr;
  var sh = sheet_(SHEETS.schedule, SCHEDULE_HEADERS);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, SCHEDULE_HEADERS.length).clearContent();
  if (list.length) {
    var rows = list.map(function (s) { return SCHEDULE_HEADERS.map(function (h) { return s[h] != null ? s[h] : ''; }); });
    sh.getRange(2, 1, rows.length, SCHEDULE_HEADERS.length).setNumberFormat('@'); // 時間欄存成文字，避免被當成時間值
    sh.getRange(2, 1, rows.length, SCHEDULE_HEADERS.length).setValues(rows);
  }
  return 'ok';
}

/* ---------- 推本週課表到 LINE（館長手動按鈕） ---------- */
function pushScheduleToLineSrv() {
  var sch = getSchedule_();
  if (!sch.length) return '課表是空的';
  var cfg = getLineConfig_();
  if (!cfg.token || !cfg.target) return '尚未設定 LINE（token／群組）';
  var order = ['週一', '週二', '週三', '週四', '週五', '週六', '週日'];
  var lines = ['🥋 本週課表'];
  order.forEach(function (wd) {
    var items = sch.filter(function (s) { return s.weekday === wd; })
      .sort(function (a, b) { return (a.start || '').localeCompare(b.start || ''); });
    if (items.length) {
      lines.push('');
      lines.push('【' + wd + '】');
      items.forEach(function (s) {
        lines.push('・' + s.classType + ' ' + (s.start || '') + '-' + (s.end || '') + (s.coach ? ' ' + s.coach : ''));
        if (s.content) lines.push('　📋 ' + s.content);
      });
    }
  });
  pushToLine_(lines.join('\n'), true);
  return 'ok';
}

/* ---------- 推今天課表到 LINE（只推今天那一天） ---------- */
function pushTodayScheduleToLineSrv() {
  var sch = getSchedule_();
  if (!sch.length) return '課表是空的';
  var cfg = getLineConfig_();
  if (!cfg.token || !cfg.target) return '尚未設定 LINE（token／群組）';
  var WD = ['週日', '週一', '週二', '週三', '週四', '週五', '週六'];
  var today = WD[new Date().getDay()];
  var items = sch.filter(function (s) { return s.weekday === today; })
    .sort(function (a, b) { return (a.start || '').localeCompare(b.start || ''); });
  if (!items.length) return '今天（' + today + '）沒有排課';
  var lines = ['🥋 今天的訓練課表（' + today + '）'];
  items.forEach(function (s) {
    lines.push('・' + s.classType + ' ' + (s.start || '') + '-' + (s.end || '') + (s.coach ? ' ' + s.coach : ''));
    if (s.content) lines.push('　📋 ' + s.content);
  });
  pushToLine_(lines.join('\n'), true);
  return 'ok';
}

/* ---------- 學生名單讀寫 ---------- */
function getStudents_() {
  var sh = sheet_(SHEETS.students, STUDENT_HEADERS);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(2, 1, last - 1, STUDENT_HEADERS.length).getValues();
  return vals.filter(function (r) { return r[1]; }).map(function (r) {
    return { id: String(r[0] || ''), name: String(r[1] || ''), classType: String(r[2] || ''), active: r[3] !== false && r[3] !== 'false' };
  });
}
function saveStudentsSrv(arr) {
  var list = (typeof arr === 'string') ? JSON.parse(arr) : arr;
  var sh = sheet_(SHEETS.students, STUDENT_HEADERS);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, STUDENT_HEADERS.length).clearContent();
  if (list.length) {
    var rows = list.map(function (s) { return [s.id || '', s.name || '', s.classType || '', s.active !== false]; });
    sh.getRange(2, 1, rows.length, STUDENT_HEADERS.length).setValues(rows);
  }
  return 'ok';
}
function writeList_(name, arr) {
  var list = (typeof arr === 'string') ? JSON.parse(arr) : arr;
  var sh = sheet_(name, ['name']);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 1).clearContent();
  if (list.length) sh.getRange(2, 1, list.length, 1).setValues(list.map(function (x) { return [x]; }));
  return 'ok';
}

/* ---------- 一次匯入（給匯入 JSON 用，可選） ---------- */
function importAllSrv(payloadJson) {
  var p = JSON.parse(payloadJson);
  if (p.coaches) writeList_(SHEETS.coaches, p.coaches);
  if (p.classes) writeList_(SHEETS.classes, p.classes);
  if (p.contents) writeList_(SHEETS.contents, p.contents);
  if (p.schedule) saveScheduleSrv(p.schedule);
  if (p.students) saveStudentsSrv(p.students);
  if (p.reviews) { (p.reviews || []).forEach(function (r) { saveReviewSrv(r); }); }
  return 'ok';
}
