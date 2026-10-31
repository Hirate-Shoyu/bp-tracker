/**
 * 血圧・脈拍・体重・服薬記録 Webアプリ（Google Apps Script / スプレッドシートにバインドして使用）
 *
 * シート「記録」  : 1行 = 1日の朝 or 夜 の血圧・脈拍
 *   A:日付 B:時間帯 C-E:1回目(最高/最低/脈拍) F-H:2回目 I-K:平均(数式) L:メモ M:更新日時
 * シート「日次」  : 1行 = 1日 の体重と服薬
 *   A:日付 B:体重 C-E:降圧剤(服用/薬品名/量) F-H:脂質異常症治療薬(服用/薬品名/量) I:更新日時
 * シート「薬の登録」: 1行 = 1つの薬の使用期間
 *   A:ID B:種類 C:薬品名 D:量 E:開始日 F:終了日(空=継続中) G:更新日時
 * 月平均と PDF 作成は Report.gs
 */

var SHEET_NAME = '記録';
var DAILY_SHEET_NAME = '日次';
var MED_SHEET_NAME = '薬の登録';
var SESSIONS = ['朝', '夜'];

// 薬の種類。key は内部用、label はシートと画面の表示用。
// oldLabels は以前の表示名。シートに残っていても読めるようにし、見つけたら label に置き換える。
var MED_TYPES = [
  { key: 'bp', label: '降圧剤', oldLabels: [] },
  { key: 'lipid', label: '脂質異常症治療薬', oldLabels: ['高脂血症の薬'] }
];
var TAKEN = '服用';
var NOT_TAKEN = '未服用';

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

var DAILY_HEADERS = ['日付', '体重(kg)'].concat(
  MED_TYPES.reduce(function (a, t) {
    return a.concat([t.label + ' 服用', t.label + ' 薬品名', t.label + ' 量']);
  }, []),
  ['更新日時']
);
var DAILY_COLS = DAILY_HEADERS.length;

var MED_HEADERS = ['ID', '種類', '薬品名', '量', '開始日', '終了日', '更新日時'];
var MED_COLS = MED_HEADERS.length;

// 入力値の許容範囲（誤入力の検出用）
var LIMITS = {
  sys: [50, 300],
  dia: [20, 200],
  pulse: [20, 250],
  weight: [20, 300]
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
  var records = readRows_(sheet, NUM_COLS).map(function (row) { return rowToRecord_(row, tz); })
    .filter(function (r) { return r; });
  records.sort(compareRecords_);
  var daily = readRows_(getDailySheet_(), DAILY_COLS).map(function (row) { return rowToDaily_(row, tz); })
    .filter(function (d) { return d; });
  daily.sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });
  return {
    records: records,
    daily: daily,
    periods: readPeriods_(tz),
    monthly: monthlyAverages_(records, tz), // Report.gs
    medTypes: MED_TYPES,
    config: CONFIG,
    today: Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'),
    nowHour: Number(Utilities.formatDate(new Date(), tz, 'H'))
  };
}

/**
 * クライアントから呼ぶ：入力画面の内容を保存する。
 * - 血圧・脈拍：1回目か2回目が入力されていれば、その日付・時間帯の行を保存（あれば上書き）
 * - 体重・服薬：日付の行を保存（あれば上書き）。その日に有効な薬の登録がない種類は記録しない
 * @param {{date:string, session:string, r1:Object, r2:Object, memo:string, weight:*, meds:Object}} input
 */
function saveRecord(input) {
  var v = validateInput_(input);
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getSheet_();
    var tz = sheet.getParent().getSpreadsheetTimeZone();
    if (v.bp) {
      var row = findRow_(sheet, v.bp.date, v.bp.session, tz);
      if (row < 0) row = sheet.getLastRow() + 1;
      sheet.getRange(row, 1, 1, NUM_COLS).setValues([recordToRow_(v.bp, tz)]);
      setAverageFormulas_(sheet, row);
      sortSheet_(sheet, NUM_COLS, [{ column: 1, ascending: true }, { column: 2, ascending: false }]);
    }
    var dailySheet = getDailySheet_();
    var drow = findRow_(dailySheet, v.date, null, tz);
    var hasDaily = v.weight !== null || MED_TYPES.some(function (t) { return v.meds[t.key] !== null; });
    // 体重・服薬の項目が送られてきたときだけ日次を更新する（空にした体重は消える）
    if (v.dailyProvided && (hasDaily || drow > 0)) {
      if (drow < 0) drow = dailySheet.getLastRow() + 1;
      var periods = readPeriods_(tz);
      dailySheet.getRange(drow, 1, 1, DAILY_COLS).setValues([dailyToRow_(v, periods, tz)]);
      sortSheet_(dailySheet, DAILY_COLS, [{ column: 1, ascending: true }]);
    }
  } finally {
    lock.releaseLock();
  }
  return getData();
}

/** クライアントから呼ぶ：指定した日付・時間帯の血圧記録を削除 */
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

/**
 * クライアントから呼ぶ：薬の使用期間を登録・更新する。
 * 新規登録で、同じ種類の「継続中」の期間が新しい開始日より前に始まっていれば、その前日で終了させる。
 * @param {{id:string, type:string, name:string, dose:string, start:string, end:string}} p
 */
function saveMedPeriod(p) {
  var period = validatePeriod_(p);
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    var periods = readPeriods_(tz);
    if (period.id) {
      var idx = indexOfId_(periods, period.id);
      if (idx < 0) throw new Error('更新対象の登録が見つかりません。画面を再読み込みしてください。');
      periods[idx] = period;
    } else {
      period.id = Utilities.getUuid();
      periods.forEach(function (q) {
        if (q.type === period.type && !q.end && q.start < period.start) q.end = addDays_(period.start, -1);
      });
      periods.push(period);
    }
    checkOverlap_(periods, period);
    writePeriods_(periods, tz);
    resyncDailyMedNames_(periods, tz);
  } finally {
    lock.releaseLock();
  }
  return getData();
}

/** クライアントから呼ぶ：薬の使用期間を削除する（日々の服用記録は残す） */
function deleteMedPeriod(id) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    var periods = readPeriods_(tz);
    var idx = indexOfId_(periods, id);
    if (idx < 0) throw new Error('削除対象の登録が見つかりません。');
    periods.splice(idx, 1);
    writePeriods_(periods, tz);
    resyncDailyMedNames_(periods, tz);
  } finally {
    lock.releaseLock();
  }
  return getData();
}

// ---------------------------------------------------------------------------
// シート

function getSheet_() {
  return ensureSheet_(SHEET_NAME, HEADERS, function (sheet) {
    sheet.getRange('A:A').setNumberFormat('yyyy-mm-dd');
    sheet.getRange('I:K').setNumberFormat('0.0');
    sheet.getRange('M:M').setNumberFormat('yyyy-mm-dd hh:mm');
  });
}

function getDailySheet_() {
  return ensureSheet_(DAILY_SHEET_NAME, DAILY_HEADERS, function (sheet) {
    sheet.getRange('A:A').setNumberFormat('yyyy-mm-dd');
    sheet.getRange('B:B').setNumberFormat('0.0#');
    var last = String.fromCharCode(64 + DAILY_COLS); // 列数は26未満
    sheet.getRange(last + ':' + last).setNumberFormat('yyyy-mm-dd hh:mm');
  });
}

function getMedSheet_() {
  var sheet = ensureSheet_(MED_SHEET_NAME, MED_HEADERS, function (sheet) {
    sheet.getRange('E:F').setNumberFormat('yyyy-mm-dd');
    sheet.getRange('G:G').setNumberFormat('yyyy-mm-dd hh:mm');
  });
  // 種類の列に以前の表示名が残っていれば、今の表示名に置き換える
  MED_TYPES.forEach(function (t) {
    t.oldLabels.forEach(function (old) {
      sheet.getRange('B:B').createTextFinder(old).matchEntireCell(true).replaceAllWith(t.label);
    });
  });
  return sheet;
}

function ensureSheet_(name, headers, format) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    format(sheet);
    return sheet;
  }
  // 見出しの表記を変えたときは、既存のシートの見出しも合わせる（列の並びは変えない前提）
  var current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  if (current.join('\t') !== headers.join('\t')) sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  return sheet;
}

function readRows_(sheet, cols) {
  var lastRow = sheet.getLastRow();
  return lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, cols).getValues() : [];
}

/** dateStr と（session を渡した場合は）時間帯が一致する行番号。なければ -1 */
function findRow_(sheet, date, session, tz) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  var values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  for (var i = 0; i < values.length; i++) {
    if (cellToDateString_(values[i][0], tz) !== date) continue;
    if (session === null || String(values[i][1]).trim() === session) return i + 2;
  }
  return -1;
}

function sortSheet_(sheet, cols, spec) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return;
  // 「記録」の時間帯は降順にすると 朝(U+671D)→夜(U+591C) の順になる
  sheet.getRange(2, 1, lastRow - 1, cols).sort(spec);
}

// ---------------------------------------------------------------------------
// 入力チェック

function validateInput_(input) {
  if (!input || !isValidDate_(input.date)) throw new Error('日付が不正です。');
  if (SESSIONS.indexOf(input.session) < 0) throw new Error('時間帯は「朝」か「夜」を指定してください。');
  var r1 = validateReading_(input.r1, '1回目');
  var r2 = validateReading_(input.r2, '2回目');
  var weight = validateWeight_(input.weight);
  var meds = {};
  MED_TYPES.forEach(function (t) {
    var m = input.meds ? input.meds[t.key] : null;
    if (m !== true && m !== false && m !== null && m !== undefined) throw new Error(t.label + 'の服用の値が不正です。');
    meds[t.key] = m === undefined ? null : m;
  });
  var bp = (r1 || r2)
    ? { date: input.date, session: input.session, r1: r1, r2: r2, memo: String(input.memo || '').slice(0, 500) }
    : null;
  var hasMeds = MED_TYPES.some(function (t) { return meds[t.key] !== null; });
  if (!bp && weight === null && !hasMeds) throw new Error('保存する内容がありません。血圧・体重・服薬のいずれかを入力してください。');
  var dailyProvided = input.weight !== undefined || input.meds !== undefined;
  return { date: input.date, bp: bp, weight: weight, meds: meds, dailyProvided: dailyProvided };
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

function validateWeight_(v) {
  if (v === null || v === undefined || String(v).trim() === '') return null;
  var s = String(v).trim();
  var n = Number(s);
  if (!/^\d+(\.\d{1,2})?$/.test(s) || !isFinite(n)) throw new Error('体重は数値（小数点以下2桁まで）で入力してください。');
  if (n < LIMITS.weight[0] || n > LIMITS.weight[1]) {
    throw new Error('体重（' + n + '）が範囲外です（' + LIMITS.weight[0] + '〜' + LIMITS.weight[1] + 'kg）。');
  }
  return n;
}

function validatePeriod_(p) {
  if (!p) throw new Error('入力がありません。');
  var type = medTypeByKey_(p.type);
  if (!type) throw new Error('薬の種類が不正です。');
  var name = String(p.name || '').trim();
  if (!name) throw new Error('薬品名を入力してください。');
  if (!isValidDate_(p.start)) throw new Error('開始日を入力してください。');
  var end = p.end ? String(p.end) : '';
  if (end && !isValidDate_(end)) throw new Error('終了日が不正です。');
  if (end && end < p.start) throw new Error('終了日が開始日より前になっています。');
  return {
    id: p.id ? String(p.id) : '',
    type: type.key,
    name: name.slice(0, 100),
    dose: String(p.dose || '').trim().slice(0, 100),
    start: p.start,
    end: end || null
  };
}

function checkOverlap_(periods, target) {
  periods.forEach(function (q) {
    if (q.id === target.id || q.type !== target.type) return;
    var qEnd = q.end || '9999-12-31', tEnd = target.end || '9999-12-31';
    if (q.start <= tEnd && target.start <= qEnd) {
      throw new Error('同じ種類の登録（' + q.name + ' ' + q.start + '〜' + (q.end || '継続中') + '）と期間が重なっています。');
    }
  });
}

// ---------------------------------------------------------------------------
// 変換

function isValidDate_(s) {
  if (typeof s !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(s)) return false;
  var p = s.split('-').map(Number);
  var d = new Date(Date.UTC(p[0], p[1] - 1, p[2]));
  return d.getUTCFullYear() === p[0] && d.getUTCMonth() === p[1] - 1 && d.getUTCDate() === p[2];
}

function addDays_(s, n) {
  var p = s.split('-').map(Number);
  return new Date(Date.UTC(p[0], p[1] - 1, p[2] + n)).toISOString().slice(0, 10);
}

function cellToDateString_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  var s = String(v || '').trim().replace(/\//g, '-');
  var m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (!m) return null;
  return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
}

function toDateCell_(s, tz) {
  return s ? Utilities.parseDate(s, tz, 'yyyy-MM-dd') : '';
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
  var cells = function (r) { return r ? [r.sys, r.dia, r.pulse] : ['', '', '']; };
  return [toDateCell_(rec.date, tz), rec.session]
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

function compareRecords_(a, b) {
  if (a.date !== b.date) return a.date < b.date ? -1 : 1;
  return SESSIONS.indexOf(a.session) - SESSIONS.indexOf(b.session);
}

function cellToTaken_(v) {
  if (v === true) return true;
  if (v === false) return false;
  var s = String(v || '').trim();
  if (s === TAKEN) return true;
  if (s === NOT_TAKEN) return false;
  return null;
}

function rowToDaily_(row, tz) {
  var date = cellToDateString_(row[0], tz);
  if (!date) return null;
  var meds = {};
  MED_TYPES.forEach(function (t, i) { meds[t.key] = cellToTaken_(row[2 + i * 3]); });
  return { date: date, weight: toNum_(row[1]), meds: meds };
}

function dailyToRow_(v, periods, tz) {
  var row = [toDateCell_(v.date, tz), v.weight === null ? '' : v.weight];
  MED_TYPES.forEach(function (t) {
    var p = activePeriod_(periods, t.key, v.date);
    var taken = p ? v.meds[t.key] : null; // 有効な登録がない日の服用は記録しない
    row.push(taken === null ? '' : taken ? TAKEN : NOT_TAKEN, p ? p.name : '', p ? p.dose : '');
  });
  row.push(new Date());
  return row;
}

// ---------------------------------------------------------------------------
// 薬の登録

function medTypeByKey_(key) {
  for (var i = 0; i < MED_TYPES.length; i++) if (MED_TYPES[i].key === key) return MED_TYPES[i];
  return null;
}

function medTypeByLabel_(label) {
  for (var i = 0; i < MED_TYPES.length; i++) {
    var t = MED_TYPES[i];
    if (t.label === label || t.oldLabels.indexOf(label) >= 0) return t;
  }
  return null;
}

function indexOfId_(periods, id) {
  for (var i = 0; i < periods.length; i++) if (periods[i].id === id) return i;
  return -1;
}

function activePeriod_(periods, typeKey, date) {
  for (var i = 0; i < periods.length; i++) {
    var p = periods[i];
    if (p.type === typeKey && p.start <= date && (!p.end || date <= p.end)) return p;
  }
  return null;
}

function readPeriods_(tz) {
  var out = [];
  readRows_(getMedSheet_(), MED_COLS).forEach(function (row) {
    var type = medTypeByLabel_(String(row[1]).trim());
    var start = cellToDateString_(row[4], tz);
    if (!type || !start || !String(row[2]).trim()) return;
    out.push({
      id: String(row[0]),
      type: type.key,
      name: String(row[2]).trim(),
      dose: String(row[3] || '').trim(),
      start: start,
      end: cellToDateString_(row[5], tz)
    });
  });
  out.sort(function (a, b) { return a.start < b.start ? -1 : a.start > b.start ? 1 : 0; });
  return out;
}

function writePeriods_(periods, tz) {
  var sheet = getMedSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) sheet.getRange(2, 1, lastRow - 1, MED_COLS).clearContent();
  if (!periods.length) return;
  var sorted = periods.slice().sort(function (a, b) {
    if (a.type !== b.type) return MED_TYPES.indexOf(medTypeByKey_(a.type)) - MED_TYPES.indexOf(medTypeByKey_(b.type));
    return a.start < b.start ? -1 : a.start > b.start ? 1 : 0;
  });
  var now = new Date();
  sheet.getRange(2, 1, sorted.length, MED_COLS).setValues(sorted.map(function (p) {
    return [p.id, medTypeByKey_(p.type).label, p.name, p.dose, toDateCell_(p.start, tz), toDateCell_(p.end, tz), now];
  }));
}

/** 登録を変えたら、「日次」シートの薬品名・量の列を登録内容に合わせて書き直す */
function resyncDailyMedNames_(periods, tz) {
  var sheet = getDailySheet_();
  var rows = readRows_(sheet, DAILY_COLS);
  if (!rows.length) return;
  rows.forEach(function (row) {
    var date = cellToDateString_(row[0], tz);
    MED_TYPES.forEach(function (t, i) {
      var p = date ? activePeriod_(periods, t.key, date) : null;
      row[3 + i * 3] = p ? p.name : '';
      row[4 + i * 3] = p ? p.dose : '';
    });
  });
  sheet.getRange(2, 1, rows.length, DAILY_COLS).setValues(rows);
}
