/**
 * 血圧・脈拍記録 Webアプリ（Google Apps Script / スプレッドシートにバインドして使用）
 *
 * シート「記録」に 1行 = 1日の朝 or 夜 として保存する。
 *   A:日付 B:時間帯 C-E:1回目(最高/最低/脈拍) F-H:2回目 I-K:平均(数式) L:メモ M:更新日時
 */

var SHEET_NAME = '記録';
var SESSIONS = ['朝', '夜'];

// グラフの参考線（mmHg）。主治医の指示に合わせて変更してください。null で非表示。
var CONFIG = {
  refSys: 135,
  refDia: 85
};

var HEADERS = [
  '日付', '時間帯',
  '1回目 最高', '1回目 最低', '1回目 脈拍',
  '2回目 最高', '2回目 最低', '2回目 脈拍',
  '平均 最高', '平均 最低', '平均 脈拍',
  'メモ', '更新日時'
];
var NUM_COLS = HEADERS.length;

// 入力値の許容範囲（誤入力の検出用）
var LIMITS = {
  sys: [50, 300],
  dia: [20, 200],
  pulse: [20, 250]
};

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('血圧記録')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

/** クライアントから呼ぶ：全記録と設定を返す */
function getData() {
  var sheet = getSheet_();
  var tz = sheet.getParent().getSpreadsheetTimeZone();
  var lastRow = sheet.getLastRow();
  var records = [];
  if (lastRow >= 2) {
    var values = sheet.getRange(2, 1, lastRow - 1, NUM_COLS).getValues();
    for (var i = 0; i < values.length; i++) {
      var rec = rowToRecord_(values[i], tz);
      if (rec) records.push(rec);
    }
  }
  records.sort(compareRecords_);
  return {
    records: records,
    config: CONFIG,
    today: Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'),
    nowHour: Number(Utilities.formatDate(new Date(), tz, 'H'))
  };
}

/**
 * クライアントから呼ぶ：1日分の朝 or 夜の記録を保存（同じ日付・時間帯があれば上書き）
 * @param {{date:string, session:string, r1:Object, r2:Object, memo:string}} input
 */
function saveRecord(input) {
  var rec = validateInput_(input);
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getSheet_();
    var tz = sheet.getParent().getSpreadsheetTimeZone();
    var row = findRow_(sheet, rec.date, rec.session, tz);
    if (row < 0) row = sheet.getLastRow() + 1;
    sheet.getRange(row, 1, 1, NUM_COLS).setValues([recordToRow_(rec, tz)]);
    setAverageFormulas_(sheet, row);
    sortSheet_(sheet);
  } finally {
    lock.releaseLock();
  }
  return getData();
}

/** クライアントから呼ぶ：指定した日付・時間帯の記録を削除 */
function deleteRecord(date, session) {
  if (!isValidDate_(date) || SESSIONS.indexOf(session) < 0) throw new Error('削除対象が不正です。');
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getSheet_();
    var row = findRow_(sheet, date, session, sheet.getParent().getSpreadsheetTimeZone());
    if (row > 0) sheet.deleteRow(row);
  } finally {
    lock.releaseLock();
  }
  return getData();
}

// ---------------------------------------------------------------------------

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, NUM_COLS).setValues([HEADERS]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.getRange('A:A').setNumberFormat('yyyy-mm-dd');
    sheet.getRange('I:K').setNumberFormat('0.0');
    sheet.getRange('M:M').setNumberFormat('yyyy-mm-dd hh:mm');
  }
  return sheet;
}

function validateInput_(input) {
  if (!input || !isValidDate_(input.date)) throw new Error('日付が不正です。');
  if (SESSIONS.indexOf(input.session) < 0) throw new Error('時間帯は「朝」か「夜」を指定してください。');
  var r1 = validateReading_(input.r1, '1回目');
  var r2 = validateReading_(input.r2, '2回目');
  if (!r1 && !r2) throw new Error('1回目か2回目のどちらかを入力してください。');
  return {
    date: input.date,
    session: input.session,
    r1: r1,
    r2: r2,
    memo: String(input.memo || '').slice(0, 500)
  };
}

/** 3項目すべて空なら null、一部だけ空ならエラー */
function validateReading_(r, label) {
  r = r || {};
  var keys = ['sys', 'dia', 'pulse'];
  var names = { sys: '最高血圧', dia: '最低血圧', pulse: '脈拍' };
  var out = {};
  var filled = 0;
  keys.forEach(function (k) {
    var v = r[k];
    if (v === null || v === undefined || v === '') return;
    var n = Number(v);
    if (!isFinite(n) || n !== Math.round(n)) throw new Error(label + 'の' + names[k] + 'は整数で入力してください。');
    if (n < LIMITS[k][0] || n > LIMITS[k][1]) {
      throw new Error(label + 'の' + names[k] + '（' + n + '）が範囲外です（' + LIMITS[k][0] + '〜' + LIMITS[k][1] + '）。');
    }
    out[k] = n;
    filled++;
  });
  if (filled === 0) return null;
  if (filled < 3) throw new Error(label + 'は最高血圧・最低血圧・脈拍の3項目を入力してください。');
  if (out.sys <= out.dia) throw new Error(label + 'の最高血圧が最低血圧以下になっています。');
  return out;
}

function isValidDate_(s) {
  if (typeof s !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(s)) return false;
  var p = s.split('-').map(Number);
  var d = new Date(Date.UTC(p[0], p[1] - 1, p[2]));
  return d.getUTCFullYear() === p[0] && d.getUTCMonth() === p[1] - 1 && d.getUTCDate() === p[2];
}

function cellToDateString_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  var s = String(v || '').trim().replace(/\//g, '-');
  var m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (!m) return null;
  return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
}

function toNum_(v) {
  if (v === '' || v === null || v === undefined) return null;
  var n = Number(v);
  return isFinite(n) ? n : null;
}

function readingFromCells_(a, b, c) {
  var r = { sys: toNum_(a), dia: toNum_(b), pulse: toNum_(c) };
  return r.sys === null && r.dia === null && r.pulse === null ? null : r;
}

function average_(readings, key) {
  var vals = readings
    .filter(function (r) { return r && r[key] !== null; })
    .map(function (r) { return r[key]; });
  if (!vals.length) return null;
  var sum = vals.reduce(function (a, b) { return a + b; }, 0);
  return Math.round((sum / vals.length) * 10) / 10;
}

function rowToRecord_(row, tz) {
  var date = cellToDateString_(row[0], tz);
  var session = String(row[1] || '').trim();
  if (!date || SESSIONS.indexOf(session) < 0) return null;
  var r1 = readingFromCells_(row[2], row[3], row[4]);
  var r2 = readingFromCells_(row[5], row[6], row[7]);
  // 平均はシートの数式に頼らず、ここで再計算する（手入力で数式が消えていても正しく出すため）
  return {
    date: date,
    session: session,
    r1: r1,
    r2: r2,
    avg: { sys: average_([r1, r2], 'sys'), dia: average_([r1, r2], 'dia'), pulse: average_([r1, r2], 'pulse') },
    memo: String(row[11] || '')
  };
}

function recordToRow_(rec, tz) {
  var dateValue = Utilities.parseDate(rec.date, tz, 'yyyy-MM-dd');
  var cells = function (r) { return r ? [r.sys, r.dia, r.pulse] : ['', '', '']; };
  return [dateValue, rec.session]
    .concat(cells(rec.r1), cells(rec.r2))
    .concat(['', '', '']) // 平均は setAverageFormulas_ で数式を入れる
    .concat([rec.memo, new Date()]);
}

function setAverageFormulas_(sheet, row) {
  // R1C1 の相対参照なので、並べ替えても各行の数式は自分の行を参照し続ける
  var f = function (c1, c2) {
    return '=IF(COUNT(RC' + c1 + ',RC' + c2 + ')=0,"",AVERAGE(RC' + c1 + ',RC' + c2 + '))';
  };
  sheet.getRange(row, 9, 1, 3).setFormulasR1C1([[f(3, 6), f(4, 7), f(5, 8)]]);
}

function findRow_(sheet, date, session, tz) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  var values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  for (var i = 0; i < values.length; i++) {
    if (cellToDateString_(values[i][0], tz) === date && String(values[i][1]).trim() === session) return i + 2;
  }
  return -1;
}

function sortSheet_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return;
  // 「朝」(U+671D) > 「夜」(U+591C) なので、時間帯は降順にすると 朝→夜 の順になる
  sheet.getRange(2, 1, lastRow - 1, NUM_COLS).sort([
    { column: 1, ascending: true },
    { column: 2, ascending: false }
  ]);
}

function compareRecords_(a, b) {
  if (a.date !== b.date) return a.date < b.date ? -1 : 1;
  return SESSIONS.indexOf(a.session) - SESSIONS.indexOf(b.session);
}
