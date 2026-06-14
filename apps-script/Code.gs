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
  students: 'Students',
  fullDb: 'FullDB',
  scheduleView: '排班課表',
  monthScheduleView: '月班表',
  lessonStatusView: '上課狀況',
  attendanceView: '出缺席紀錄',
  kpiView: 'KPI復盤',
  courseRecords: '課程紀錄表',
  studentRecords: '學生個別紀錄表',
  coachRecords: '教練資料表',
  classRecords: '班級資料表',
  statsReport: '統計報表表',
  parentReports: '家長回報紀錄表'
};
var SCHEDULE_HEADERS = ['id', 'weekday', 'classType', 'start', 'end', 'coach', 'content'];
var STUDENT_HEADERS = ['id', 'name', 'classType', 'active'];
var FULLDB_HEADERS = ['key', 'updatedAt', 'json'];
var SCHEDULE_VIEW_HEADERS = ['日期','星期','班別','開始','結束','主教練','助教1','助教2','助教3','上課內容','狀態','備註'];
var MONTH_SCHEDULE_HEADERS = ['月份','日期','星期','班別','開始','結束','主教練','助教','上課內容','狀態','備註'];
var LESSON_STATUS_HEADERS = ['日期','班別','開始','結束','人員','職務','上課狀態','已完成','遲到','代班','完成時間','課程狀態'];
var ATTENDANCE_HEADERS = ['日期','星期','班別','開始','結束','人員','職務','出缺席','是否遲到','是否代班','課程狀態','備註'];
var COURSE_RECORD_HEADERS = ['日期','班別','上課教練','助教','今日課程主題','學生人數','上課內容','整體表現','專注度','秩序','體能','技術完成度','總分','今日亮點','需要改善','下次課程建議','受傷或特殊狀況','家長回報'];
var STUDENT_RECORD_HEADERS = ['日期','班別','教練','學生姓名','今日表現','技術問題','態度問題','專注問題','下次提醒事項','需要通知家長','教練備註'];
var COACH_RECORD_HEADERS = ['教練','角色','啟用','本月上課堂數','本月復盤數','復盤完成率','平均課程評分','家長回報次數','館長備註','酬勞參考'];
var CLASS_RECORD_HEADERS = ['班級','預設開始','預設結束','預設主教練','預設助教人數','本月復盤數','平均分數','備註'];
var STATS_REPORT_HEADERS = ['項目','數值','說明'];
var PARENT_REPORT_HEADERS = ['日期','班別','教練','是否已發送','LINE文字'];
var KPI_VIEW_HEADERS = [
  '日期','教練','班別','學生人數','上課內容','總分','等級',
  '教學設計','技術示範','班級控場','學生互動','安全管理','課後延續','課程品質',
  '課前準備','準時到課','上課態度','課後反思',
  '課程目標','備課設計','觀察重點','上次延續','最成功','待改善','具體證據','下次改善','下次備課','追蹤指標','協助需求','建立時間'
];

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

/* ---------- 完整系統 DB 同步（給新版排班/薪資/復盤系統使用） ---------- */
function getFullDbSrv() {
  var sh = sheet_(SHEETS.fullDb, FULLDB_HEADERS);
  var last = sh.getLastRow();
  if (last < 2) return JSON.stringify({ updatedAt: 0, db: null });
  var vals = sh.getRange(2, 1, last - 1, FULLDB_HEADERS.length).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (vals[i][0] === 'hcps_db_v1') {
      return JSON.stringify({
        updatedAt: Number(vals[i][1] || 0),
        db: vals[i][2] ? JSON.parse(vals[i][2]) : null
      });
    }
  }
  return JSON.stringify({ updatedAt: 0, db: null });
}

function saveFullDbSrv(payloadJson) {
  var payload = (typeof payloadJson === 'string') ? JSON.parse(payloadJson) : payloadJson;
  var sh = sheet_(SHEETS.fullDb, FULLDB_HEADERS);
  var updatedAt = Number(payload.updatedAt || new Date().getTime());
  var db = payload.db || {};
  var dbJson = JSON.stringify(db);
  var last = sh.getLastRow();
  var found = 0;
  if (last >= 2) {
    var keys = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < keys.length; i++) {
      if (keys[i][0] === 'hcps_db_v1') { found = i + 2; break; }
    }
  }
  var row = ['hcps_db_v1', updatedAt, dbJson];
  if (found) sh.getRange(found, 1, 1, FULLDB_HEADERS.length).setValues([row]);
  else sh.appendRow(row);
  syncReadableSheets_(db);
  return JSON.stringify({ ok: true, updatedAt: updatedAt });
}

function rebuildReadableSheetsSrv() {
  var raw = JSON.parse(getFullDbSrv());
  if (!raw.db) return 'FullDB 尚無資料，請先從系統儲存一次';
  syncReadableSheets_(raw.db);
  return 'ok';
}

function resetSheet_(name, headers) {
  var sh = sheet_(name, headers);
  sh.clearContents();
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  return sh;
}

function writeRows_(name, headers, rows) {
  var sh = resetSheet_(name, headers);
  if (rows.length) sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  sh.autoResizeColumns(1, headers.length);
}

function syncReadableSheets_(db) {
  db = db || {};
  writeScheduleView_(db);
  writeMonthScheduleView_(db);
  writeLessonStatusView_(db);
  writeAttendanceView_(db);
  writeKpiView_(db);
  writeCourseRecords_(db);
  writeStudentRecords_(db);
  writeCoachRecords_(db);
  writeClassRecords_(db);
  writeStatsReport_(db);
  writeParentReports_(db);
}

function scheduleEntries_(s) {
  if (s.entries && s.entries.length) return s.entries;
  var out = [];
  if (s.mainCoach) out.push({ person: s.mainCoach, slot: '主教練', status: s.status || '已排班' });
  ['a1','a2','a3'].forEach(function (k, idx) {
    if (s[k]) out.push({ person: s[k], slot: '助教' + (idx + 1), status: s.status || '已排班' });
  });
  return out;
}

function writeScheduleView_(db) {
  var rows = (db.schedules || []).map(function (s) {
    return [
      s.date || '', s.weekday || '', s.className || '', s.start || '', s.end || '',
      s.mainCoach || '', s.a1 || '', s.a2 || '', s.a3 || '',
      s.content || '', s.status || '', s.note || ''
    ];
  });
  writeRows_(SHEETS.scheduleView, SCHEDULE_VIEW_HEADERS, rows);
}

function writeMonthScheduleView_(db) {
  var rows = (db.schedules || [])
    .slice()
    .sort(function (a, b) { return ((a.date || '') + (a.start || '')).localeCompare((b.date || '') + (b.start || '')); })
    .map(function (s) {
      return [
        (s.date || '').slice(0, 7), s.date || '', s.weekday || '', s.className || '',
        s.start || '', s.end || '', s.mainCoach || '',
        [s.a1, s.a2, s.a3].filter(Boolean).join('、'),
        s.content || '', s.status || '', s.note || ''
      ];
    });
  writeRows_(SHEETS.monthScheduleView, MONTH_SCHEDULE_HEADERS, rows);
}

function writeLessonStatusView_(db) {
  var rows = [];
  (db.schedules || []).forEach(function (s) {
    scheduleEntries_(s).forEach(function (e) {
      rows.push([
        s.date || '', s.className || '', s.start || '', s.end || '',
        e.person || '', e.slot || '', e.status || '',
        e.completed ? 'Y' : '', e.late ? 'Y' : '', e.isSub ? 'Y' : '',
        e.completedAt || '', s.status || ''
      ]);
    });
  });
  writeRows_(SHEETS.lessonStatusView, LESSON_STATUS_HEADERS, rows);
}

function attendanceLabel_(s, e) {
  if ((s.status || '') === '取消' || (e.status || '') === '取消') return '取消';
  if ((e.status || '') === '請假' || (s.status || '') === '請假') return '請假';
  if (e.completed && e.late) return '遲到';
  if (e.completed) return '出席';
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  if ((s.date || '') < today) return '缺席';
  return '未到課';
}

function writeAttendanceView_(db) {
  var rows = [];
  (db.schedules || []).forEach(function (s) {
    scheduleEntries_(s).forEach(function (e) {
      rows.push([
        s.date || '', s.weekday || '', s.className || '', s.start || '', s.end || '',
        e.person || '', e.slot || '', attendanceLabel_(s, e),
        e.late ? 'Y' : '', e.isSub ? 'Y' : '', s.status || '', s.note || ''
      ]);
    });
  });
  writeRows_(SHEETS.attendanceView, ATTENDANCE_HEADERS, rows);
}

function writeKpiView_(db) {
  var rows = (db.reviews || []).map(function (r) {
    var scores = r.scores || [];
    var self = r.self || {};
    var notes = r.selfNotes || {};
    return [
      r.date || '', r.coach || '', r.className || '', r.count || '', r.content || '',
      r.total || '', r.grade || '',
      scores[0] || '', scores[1] || '', scores[2] || '', scores[3] || '', scores[4] || '', scores[5] || '',
      r.quality || '',
      self.prep || '', self.punctual || '', self.attitude || '', self.reflection || '',
      r.goal || '', r.plan || '', r.focus || '', r.follow || '',
      r.best || '', r.worst || '', r.evidence || '', r.next || '', r.nextPlan || '', r.metric || '', r.support || '',
      r.ts ? new Date(r.ts) : ''
    ];
  });
  writeRows_(SHEETS.kpiView, KPI_VIEW_HEADERS, rows);
}

function monthKey_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
}

function completedLessonsForCoach_(db, name, month) {
  var count = 0;
  (db.schedules || []).forEach(function (s) {
    if ((s.date || '').slice(0, 7) !== month) return;
    scheduleEntries_(s).forEach(function (e) {
      if (e.person === name && e.completed) count++;
    });
  });
  return count;
}

function writeCourseRecords_(db) {
  var rows = (db.reviews || []).map(function (r) {
    var q = r.quick || {};
    return [
      r.date || '', r.className || '', r.coach || '', r.assistant || '',
      (r.topics || []).join('、'), r.count || '', r.content || '',
      q.overall || '', q.focus || '', q.order || '', q.fitness || '', q.technique || '',
      r.total || '', r.best || '', r.worst || '', r.courseNext || r.next || '', r.special || '', r.parentReport || ''
    ];
  });
  writeRows_(SHEETS.courseRecords, COURSE_RECORD_HEADERS, rows);
}

function writeStudentRecords_(db) {
  var rows = (db.studentRecords || []).map(function (r) {
    return [
      r.date || '', r.className || '', r.coach || '', r.studentName || '', r.performance || '',
      r.techIssue || '', r.attitudeIssue || '', r.focusIssue || '', r.nextReminder || '',
      r.notifyParent ? 'Y' : '', r.note || ''
    ];
  });
  writeRows_(SHEETS.studentRecords, STUDENT_RECORD_HEADERS, rows);
}

function writeCoachRecords_(db) {
  var month = monthKey_();
  var rows = (db.people || []).map(function (p) {
    var reviews = (db.reviews || []).filter(function (r) { return r.coach === p.name && (r.date || '').slice(0, 7) === month; });
    var lessons = completedLessonsForCoach_(db, p.name, month);
    var avg = reviews.length ? reviews.reduce(function (a, r) { return a + (Number(r.total) || 0); }, 0) / reviews.length : 0;
    var rate = lessons ? Math.round(reviews.length / lessons * 100) : 0;
    var parentCount = (db.parentReports || []).filter(function (r) { return r.coach === p.name && (r.date || '').slice(0, 7) === month; }).length;
    return [
      p.name || '', p.role || '', p.active !== false ? 'Y' : '',
      lessons, reviews.length, rate + '%', avg ? avg.toFixed(1) : '',
      parentCount, p.note || '', (rate >= 80 && avg >= 3.5) ? '可參考' : '暫不建議'
    ];
  });
  writeRows_(SHEETS.coachRecords, COACH_RECORD_HEADERS, rows);
}

function writeClassRecords_(db) {
  var month = monthKey_();
  var rows = (db.classes || []).map(function (c) {
    var reviews = (db.reviews || []).filter(function (r) { return r.className === c.name && (r.date || '').slice(0, 7) === month; });
    var avg = reviews.length ? reviews.reduce(function (a, r) { return a + (Number(r.total) || 0); }, 0) / reviews.length : 0;
    return [c.name || '', c.start || '', c.end || '', c.mainCoach || '', c.assistantCount || '', reviews.length, avg ? avg.toFixed(1) : '', c.note || ''];
  });
  writeRows_(SHEETS.classRecords, CLASS_RECORD_HEADERS, rows);
}

function writeStatsReport_(db) {
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var month = monthKey_();
  var reviews = db.reviews || [];
  var todayReviews = reviews.filter(function (r) { return r.date === today; });
  var monthReviews = reviews.filter(function (r) { return (r.date || '').slice(0, 7) === month; });
  var avg = monthReviews.length ? monthReviews.reduce(function (a, r) { return a + (Number(r.total) || 0); }, 0) / monthReviews.length : 0;
  var concern = (db.studentRecords || []).filter(function (r) { return r.performance === '需提醒' || r.notifyParent; }).length;
  var parentCount = (db.parentReports || []).filter(function (r) { return (r.date || '').slice(0, 7) === month; }).length;
  var rows = [
    ['今日完成復盤堂數', todayReviews.length, today],
    ['本月復盤堂數', monthReviews.length, month],
    ['本月平均課程評分', avg ? avg.toFixed(1) : '', month],
    ['學生需關注筆數', concern, '今日表現需提醒或需通知家長'],
    ['家長回報紀錄數', parentCount, month]
  ];
  writeRows_(SHEETS.statsReport, STATS_REPORT_HEADERS, rows);
}

function writeParentReports_(db) {
  var rows = (db.parentReports || []).map(function (r) {
    return [r.date || '', r.className || '', r.coach || '', r.sent ? 'Y' : '', r.text || ''];
  });
  writeRows_(SHEETS.parentReports, PARENT_REPORT_HEADERS, rows);
}
