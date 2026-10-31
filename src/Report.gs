/**
 * 月平均の計算と PDF 作成（A4 縦）
 *
 * シート「月平均」: 過去の月平均値（取り込み用）。1行 = 1か月
 *   A:年月（2015-11、2015/11、201511 など） B-D:朝(最高/最低/脈拍) E-G:夜(最高/最低/脈拍)
 * アプリの記録から計算した月平均と合わせて使う。同じ月があれば「月平均」シートの値を優先する。
 *
 * PDF は Google ドキュメントを A4 で組み立てて変換し、ドライブのフォルダ「血圧記録 PDF」に保存する。
 * グラフ画像は画面側（Chart.js）で作った PNG を受け取って貼り込む。
 */

var MONTHLY_SHEET_NAME = '月平均';
var MONTHLY_HEADERS = ['年月', '朝 最高', '朝 最低', '朝 脈拍', '夜 最高', '夜 最低', '夜 脈拍'];
var MONTHLY_COLS = MONTHLY_HEADERS.length;
var PDF_FOLDER_NAME = '血圧記録 PDF';
var PDF_FOLDER_PROP = 'pdfFolderId';

// A4 縦（ポイント）
var A4 = { width: 595.28, height: 841.89, margin: 36 };
var CONTENT_WIDTH = A4.width - A4.margin * 2;

var SESSION_KEYS = { '朝': 'morning', '夜': 'night' };
var WEEKDAYS = ['日', '月', '火', '水', '木', '金', '土'];

// ---------------------------------------------------------------------------
// 月平均

function getMonthlySheet_() {
  return ensureSheet_(MONTHLY_SHEET_NAME, MONTHLY_HEADERS, function (sheet) {
    sheet.getRange('A:A').setNumberFormat('@'); // 「2015-01」を日付に変換させない
  });
}

/** 各日の2回平均を、朝・夜それぞれ月ごとに平均する */
function computeMonthly_(records) {
  var acc = {};
  records.forEach(function (r) {
    var m = r.date.slice(0, 7);
    var key = SESSION_KEYS[r.session];
    acc[m] = acc[m] || { morning: [], night: [] };
    acc[m][key].push(r.avg);
  });
  var out = {};
  Object.keys(acc).forEach(function (m) {
    out[m] = { month: m, source: 'app', morning: sessionAverage_(acc[m].morning), night: sessionAverage_(acc[m].night) };
  });
  return out;
}

function sessionAverage_(avgs) {
  if (!avgs.length) return null;
  var mean = function (k) {
    var v = avgs.map(function (a) { return a[k]; }).filter(function (x) { return x !== null; });
    if (!v.length) return null;
    return Math.round(v.reduce(function (a, b) { return a + b; }, 0) / v.length * 10) / 10;
  };
  return { sys: mean('sys'), dia: mean('dia'), pulse: mean('pulse'), days: avgs.length };
}

function cellToMonth_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM');
  var s = String(v === null || v === undefined ? '' : v).trim();
  var m = s.match(/^(\d{4})\s*[-\/年.]\s*(\d{1,2})/) || s.match(/^(\d{4})(\d{2})$/); // 「201511」形式も可
  if (!m || Number(m[2]) < 1 || Number(m[2]) > 12) return null;
  return m[1] + '-' + ('0' + m[2]).slice(-2);
}

function readImportedMonthly_(tz) {
  var out = [];
  readRows_(getMonthlySheet_(), MONTHLY_COLS).forEach(function (row) {
    var month = cellToMonth_(row[0], tz);
    if (!month) return;
    var s = function (i) {
      var r = { sys: toNum_(row[i]), dia: toNum_(row[i + 1]), pulse: toNum_(row[i + 2]), days: null };
      return r.sys === null && r.dia === null && r.pulse === null ? null : r;
    };
    var rec = { month: month, source: 'import', morning: s(1), night: s(4) };
    if (rec.morning || rec.night) out.push(rec);
  });
  return out;
}

/** 取り込み値とアプリの記録からの計算値を合わせた月平均（年月の昇順） */
function monthlyAverages_(records, tz) {
  var all = computeMonthly_(records);
  readImportedMonthly_(tz).forEach(function (m) { all[m.month] = m; });
  return Object.keys(all).sort().map(function (k) { return all[k]; });
}

// ---------------------------------------------------------------------------
// 月ごとの PDF

/**
 * クライアントから呼ぶ：指定月の PDF を作る
 * 1枚目：測定値の表 / 2枚目：月平均値とグラフ
 * @param {string} month 'yyyy-MM'
 * @param {{bp:string, pulse:string}} charts グラフの PNG（data URL）
 */
function createMonthlyReportPdf(month, charts) {
  if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(String(month))) throw new Error('年月が不正です。');
  var data = getData();
  var recs = data.records.filter(function (r) { return r.date.slice(0, 7) === month; });
  if (!recs.length) throw new Error(monthLabel_(month) + 'の血圧の記録がありません。');
  var avg = computeMonthly_(recs)[month];
  var title = monthLabel_(month);

  return buildPdf_('血圧記録_' + month, function (body) {
    // 1枚目：測定値
    heading_(body, '血圧・脈拍 測定記録　' + title, true);
    var table = body.appendTable(monthTableRows_(month, recs));
    styleMonthTable_(table);
    note_(body, '単位：血圧 mmHg、脈拍 回/分。各時間帯に2回測定。');
    body.appendPageBreak();

    // 2枚目：月平均とグラフ
    heading_(body, '月平均・グラフ　' + title, false);
    var sum = body.appendTable(monthAverageRows_(avg));
    styleSimpleTable_(sum, [80, 80, 80, 80, 80]);
    note_(body, '月平均：各日の2回の平均値を、さらに月で平均した値です。');
    subheading_(body, '血圧（朝・夜それぞれ2回の平均, mmHg）');
    appendChart_(body, charts && charts.bp);
    subheading_(body, '脈拍（朝・夜それぞれ2回の平均, 回/分）');
    appendChart_(body, charts && charts.pulse);
  });
}

function monthLabel_(month) {
  return Number(month.slice(0, 4)) + '年' + Number(month.slice(5, 7)) + '月';
}

function daysInMonth_(month) {
  var y = Number(month.slice(0, 4)), m = Number(month.slice(5, 7));
  return new Date(Date.UTC(y, m, 0)).getUTCDate();
}

/** 1枚目の表：見出し2行 + その月の全日（記録のない日は空欄） */
function monthTableRows_(month, recs) {
  var byKey = {};
  recs.forEach(function (r) { byKey[r.date + r.session] = r; });
  var group = ['朝1回目', '朝2回目', '夜1回目', '夜2回目'];
  var head1 = ['', ''], head2 = ['日', '曜'];
  group.forEach(function (g) {
    head1.push('', g, '');
    head2.push('最高', '最低', '脈拍');
  });
  var rows = [head1, head2];
  var n = daysInMonth_(month);
  for (var d = 1; d <= n; d++) {
    var date = month + '-' + ('0' + d).slice(-2);
    var row = [String(d), WEEKDAYS[new Date(Date.UTC(Number(month.slice(0, 4)), Number(month.slice(5, 7)) - 1, d)).getUTCDay()]];
    SESSIONS.forEach(function (s) {
      var r = byKey[date + s];
      [r && r.r1, r && r.r2].forEach(function (x) {
        row.push(x ? numText_(x.sys) : '', x ? numText_(x.dia) : '', x ? numText_(x.pulse) : '');
      });
    });
    rows.push(row);
  }
  return rows;
}

function monthAverageRows_(avg) {
  var rows = [['', '測定日数', '最高', '最低', '脈拍']];
  [['朝', avg.morning], ['夜', avg.night]].forEach(function (p) {
    var a = p[1];
    rows.push([p[0], a ? String(a.days) : '0', a ? numText_(a.sys, 1) : '—', a ? numText_(a.dia, 1) : '—', a ? numText_(a.pulse, 1) : '—']);
  });
  return rows;
}

function numText_(v, digits) {
  if (v === null || v === undefined) return digits ? '—' : '';
  return digits ? (Math.round(v * 10) / 10).toFixed(digits) : String(v);
}

// ---------------------------------------------------------------------------
// 月平均グラフの PDF

/**
 * クライアントから呼ぶ：月平均グラフ（血圧）の PDF を作る。1ページに上下2段。
 * @param {{title:string, image:string}[]} panels 1段 = 3年分のグラフ
 * @param {string} range ファイル名用の期間（例 '2015-2026'）
 */
function createMonthlyChartPdf(panels, range) {
  if (!panels || !panels.length) throw new Error('グラフのデータがありません。');
  if (!/^\d{4}-\d{4}$/.test(String(range))) throw new Error('期間が不正です。');
  return buildPdf_('血圧月平均グラフ_' + range, function (body) {
    panels.forEach(function (p, i) {
      if (i > 0 && i % 2 === 0) body.appendPageBreak();
      heading_(body, String(p.title).slice(0, 100), i === 0);
      appendChart_(body, p.image);
    });
  });
}

// ---------------------------------------------------------------------------
// ドキュメント → PDF

function buildPdf_(name, fill) {
  var doc = DocumentApp.create(name + '（作成中）');
  var tempFile = DriveApp.getFileById(doc.getId());
  try {
    var body = doc.getBody();
    body.setPageWidth(A4.width).setPageHeight(A4.height)
      .setMarginTop(A4.margin).setMarginBottom(A4.margin)
      .setMarginLeft(A4.margin).setMarginRight(A4.margin);
    fill(body);
    doc.saveAndClose();
    var pdf = tempFile.getAs('application/pdf').setName(name + '.pdf');
    var folder = getPdfFolder_();
    var old = folder.getFilesByName(name + '.pdf');
    while (old.hasNext()) old.next().setTrashed(true); // 同じ名前の古い PDF はゴミ箱へ
    var file = folder.createFile(pdf);
    return { name: file.getName(), url: file.getUrl() };
  } finally {
    tempFile.setTrashed(true);
  }
}

function getPdfFolder_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty(PDF_FOLDER_PROP);
  if (id) {
    try {
      var f = DriveApp.getFolderById(id);
      if (!f.isTrashed()) return f;
    } catch (e) { /* 削除済みなら作り直す */ }
  }
  var folder = DriveApp.createFolder(PDF_FOLDER_NAME);
  props.setProperty(PDF_FOLDER_PROP, folder.getId());
  return folder;
}

/** 最初の見出しは文書冒頭の空段落を使う（先頭に空行が入らないように） */
function heading_(body, text, first) {
  var p = first ? body.getChild(0).asParagraph() : body.appendParagraph('');
  p.setText(text);
  p.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
  p.setSpacingBefore(0).setSpacingAfter(6).setLineSpacing(1);
  p.editAsText().setFontSize(14).setBold(true);
  return p;
}

function subheading_(body, text) {
  var p = body.appendParagraph(text);
  p.setSpacingBefore(8).setSpacingAfter(2).setLineSpacing(1);
  p.editAsText().setFontSize(10).setBold(true);
}

function note_(body, text) {
  var p = body.appendParagraph(text);
  p.setSpacingBefore(2).setSpacingAfter(0).setLineSpacing(1);
  p.editAsText().setFontSize(8).setBold(false).setForegroundColor('#555555');
}

/** 本文の幅いっぱいに、画像の縦横比を保って貼る */
function appendChart_(body, dataUrl) {
  var m = String(dataUrl || '').match(/^data:image\/png;base64,(.+)$/);
  if (!m) throw new Error('グラフ画像がありません。');
  var blob = Utilities.newBlob(Utilities.base64Decode(m[1]), 'image/png', 'chart.png');
  var p = body.appendParagraph('');
  p.setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1);
  var img = p.appendInlineImage(blob);
  var ratio = img.getHeight() / img.getWidth();
  img.setWidth(CONTENT_WIDTH).setHeight(Math.round(CONTENT_WIDTH * ratio));
}

function styleCells_(table, fontSize, padding) {
  for (var i = 0; i < table.getNumRows(); i++) {
    var row = table.getRow(i);
    for (var j = 0; j < row.getNumCells(); j++) {
      var cell = row.getCell(j);
      cell.setPaddingTop(padding).setPaddingBottom(padding).setPaddingLeft(2).setPaddingRight(2);
      cell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
      var p = cell.getChild(0).asParagraph();
      p.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1);
      p.editAsText().setFontSize(fontSize);
    }
  }
}

function styleMonthTable_(table) {
  table.setBorderWidth(0.5).setBorderColor('#999999');
  table.setColumnWidth(0, 24).setColumnWidth(1, 24);
  var w = (CONTENT_WIDTH - 48) / 12;
  for (var c = 2; c < 14; c++) table.setColumnWidth(c, w);
  styleCells_(table, 9, 2.5);
  var tint = function (c) { return c < 8 ? '#e3eefb' : '#fdeadf'; }; // 朝＝青系、夜＝橙系
  for (var r = 0; r < 2; r++) {
    for (var c2 = 0; c2 < 14; c2++) {
      var cell = table.getRow(r).getCell(c2);
      cell.setBackgroundColor(c2 < 2 ? '#eeeeee' : tint(c2));
      cell.editAsText().setBold(true).setFontSize(8);
    }
  }
  // 土日の行は薄く色付け
  for (var i = 2; i < table.getNumRows(); i++) {
    var wd = table.getRow(i).getCell(1).getText();
    if (wd === '土' || wd === '日') {
      for (var k = 0; k < 14; k++) table.getRow(i).getCell(k).setBackgroundColor('#f5f5f5');
    }
  }
}

function styleSimpleTable_(table, widths) {
  table.setBorderWidth(0.5).setBorderColor('#999999');
  widths.forEach(function (w, i) { table.setColumnWidth(i, w); });
  styleCells_(table, 10, 3);
  for (var c = 0; c < widths.length; c++) {
    table.getRow(0).getCell(c).setBackgroundColor('#eeeeee').editAsText().setBold(true);
  }
}
