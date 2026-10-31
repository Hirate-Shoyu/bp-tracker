// Code.gs のロジックをモック上で検証する： node --test test/
const test = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

function load() {
  const ctx = vm.createContext({ console, Date, Math, Number, String, isFinite, Error, Array, Object, Proxy, encodeURIComponent });
  vm.runInContext(fs.readFileSync(path.join(__dirname, 'mock-gas.js'), 'utf8'), ctx);
  vm.runInContext(fs.readFileSync(path.join(__dirname, '../src/Code.gs'), 'utf8'), ctx);
  vm.runInContext(fs.readFileSync(path.join(__dirname, '../src/Report.gs'), 'utf8'), ctx);
  return ctx;
}
const plain = (o) => JSON.parse(JSON.stringify(o));

test('保存すると1行=日付×時間帯で記録され、平均が計算される', () => {
  const g = load();
  const d = g.saveRecord({ date: '2026-10-30', session: '朝', r1: { sys: 130, dia: 85, pulse: 70 }, r2: { sys: 126, dia: 80, pulse: 66 }, memo: '' });
  assert.strictEqual(d.records.length, 1);
  assert.deepStrictEqual(plain(d.records[0].avg), { sys: 128, dia: 82.5, pulse: 68 });
  const sheet = g.__mockSpreadsheet.getSheetByName('記録');
  assert.strictEqual(sheet.rows[0][0], '日付');
  assert.strictEqual(sheet.rows[1][8], '=IF(COUNT(RC3,RC6)=0,"",AVERAGE(RC3,RC6))');
});

test('同じ日付・時間帯は上書き、朝夜と日付順に並ぶ', () => {
  const g = load();
  const r = (sys, dia, pulse) => ({ sys, dia, pulse });
  g.saveRecord({ date: '2026-10-31', session: '夜', r1: r(120, 80, 60), r2: {} });
  g.saveRecord({ date: '2026-10-30', session: '夜', r1: r(121, 81, 61), r2: {} });
  g.saveRecord({ date: '2026-10-31', session: '朝', r1: r(140, 90, 70), r2: r(138, 88, 72) });
  const d = g.saveRecord({ date: '2026-10-31', session: '夜', r1: r(118, 78, 62), r2: r(116, 76, 60) });
  assert.deepStrictEqual(plain(d.records.map((x) => x.date + x.session)), ['2026-10-30夜', '2026-10-31朝', '2026-10-31夜']);
  assert.strictEqual(d.records[2].avg.sys, 117);
  const sheet = g.__mockSpreadsheet.getSheetByName('記録');
  assert.deepStrictEqual(plain(sheet.rows.slice(1).map((x) => x[1])), ['夜', '朝', '夜']);
  assert.strictEqual(sheet.getLastRow(), 4);
});

test('1回分だけでも保存でき、平均はその値', () => {
  const g = load();
  const d = g.saveRecord({ date: '2026-10-31', session: '朝', r1: {}, r2: { sys: '125', dia: '79', pulse: '64' } });
  assert.strictEqual(d.records[0].r1, null);
  assert.deepStrictEqual(plain(d.records[0].avg), { sys: 125, dia: 79, pulse: 64 });
});

test('入力チェック', () => {
  const g = load();
  const ok = { sys: 120, dia: 80, pulse: 60 };
  assert.throws(() => g.saveRecord({ date: '2026-02-30', session: '朝', r1: ok }), /日付/);
  assert.throws(() => g.saveRecord({ date: '2026-10-31', session: '昼', r1: ok }), /時間帯/);
  assert.throws(() => g.saveRecord({ date: '2026-10-31', session: '朝', r1: {}, r2: {} }), /保存する内容/);
  assert.throws(() => g.saveRecord({ date: '2026-10-31', session: '朝', r1: { sys: 120, dia: 80 } }), /3項目/);
  assert.throws(() => g.saveRecord({ date: '2026-10-31', session: '朝', r1: { sys: 80, dia: 120, pulse: 60 } }), /最低血圧以下/);
  assert.throws(() => g.saveRecord({ date: '2026-10-31', session: '朝', r1: { sys: 1200, dia: 80, pulse: 60 } }), /範囲外/);
  assert.throws(() => g.saveRecord({ date: '2026-10-31', session: '朝', r1: { sys: 120.5, dia: 80, pulse: 60 } }), /整数/);
});

test('削除', () => {
  const g = load();
  g.saveRecord({ date: '2026-10-31', session: '朝', r1: { sys: 120, dia: 80, pulse: 60 } });
  g.saveRecord({ date: '2026-10-31', session: '夜', r1: { sys: 118, dia: 78, pulse: 58 } });
  const d = g.deleteRecord('2026-10-31', '朝');
  assert.deepStrictEqual(plain(d.records.map((x) => x.session)), ['夜']);
});

test('シートに手入力された文字列の日付も読める', () => {
  const g = load();
  const sheet = g.getSheet_();
  sheet.getRange(2, 1, 1, 13).setValues([['2026/10/5', '朝', 130, 80, 70, '', '', '', '', '', '', '', '']]);
  const d = g.getData();
  assert.strictEqual(d.records[0].date, '2026-10-05');
  assert.strictEqual(d.records[0].avg.sys, 130);
});

// ---- 体重・服薬 ---------------------------------------------------------

const BP = { sys: 120, dia: 80, pulse: 60 };
const period = (type, name, dose, start, end) => ({ type, name, dose, start, end: end || '' });

test('薬の登録：新規登録で継続中の前の期間が前日で終わる', () => {
  const g = load();
  g.saveMedPeriod(period('bp', 'アムロジピン', '5mg', '2026-09-01'));
  g.saveMedPeriod(period('lipid', 'ロスバスタチン', '2.5mg', '2026-09-01'));
  const d = g.saveMedPeriod(period('bp', 'アムロジピン', '10mg', '2026-10-15'));
  const bp = plain(d.periods.filter((p) => p.type === 'bp').map((p) => [p.dose, p.start, p.end]));
  assert.deepStrictEqual(bp, [['5mg', '2026-09-01', '2026-10-14'], ['10mg', '2026-10-15', null]]);
  const rows = g.__mockSpreadsheet.getSheetByName('薬の登録').rows.slice(1).map((r) => r[1] + r[3]);
  assert.deepStrictEqual(plain(rows), ['降圧剤5mg', '降圧剤10mg', '高脂血症の薬2.5mg']);
});

test('薬の登録：期間の重なり・入力不備はエラー、編集と削除ができる', () => {
  const g = load();
  let d = g.saveMedPeriod(period('bp', 'A', '5mg', '2026-09-01', '2026-09-30'));
  assert.throws(() => g.saveMedPeriod(period('bp', 'B', '5mg', '2026-09-15', '2026-10-10')), /重なって/);
  assert.throws(() => g.saveMedPeriod(period('bp', '', '5mg', '2026-10-01')), /薬品名/);
  assert.throws(() => g.saveMedPeriod(period('bp', 'B', '5mg', '2026-10-05', '2026-10-01')), /終了日/);
  assert.throws(() => g.saveMedPeriod(period('xx', 'B', '5mg', '2026-10-05')), /種類/);
  const id = d.periods[0].id;
  d = g.saveMedPeriod({ id, type: 'bp', name: 'A', dose: '2.5mg', start: '2026-09-01', end: '' });
  assert.strictEqual(d.periods.length, 1);
  assert.strictEqual(d.periods[0].dose, '2.5mg');
  assert.strictEqual(d.periods[0].end, null);
  d = g.deleteMedPeriod(id);
  assert.strictEqual(d.periods.length, 0);
});

test('日次：体重と服薬を保存。有効な登録がない薬は記録しない', () => {
  const g = load();
  g.saveMedPeriod(period('bp', 'アムロジピン', '5mg', '2026-10-01'));
  let d = g.saveRecord({ date: '2026-10-31', session: '朝', r1: BP, weight: '65.4', meds: { bp: true, lipid: true } });
  assert.strictEqual(d.records.length, 1);
  assert.deepStrictEqual(plain(d.daily), [{ date: '2026-10-31', weight: 65.4, meds: { bp: true, lipid: null } }]);
  const row = g.__mockSpreadsheet.getSheetByName('日次').rows[1];
  assert.deepStrictEqual(plain(row.slice(1, 8)), [65.4, '服用', 'アムロジピン', '5mg', '', '', '']);
  // 夜に同じ日を保存すると日次は上書き（体重を空にすると消える）
  d = g.saveRecord({ date: '2026-10-31', session: '夜', r1: BP, weight: '', meds: { bp: false } });
  assert.deepStrictEqual(plain(d.daily), [{ date: '2026-10-31', weight: null, meds: { bp: false, lipid: null } }]);
  assert.strictEqual(d.records.length, 2);
});

test('日次：血圧なしで体重・服薬だけ保存できる。何もなければエラー', () => {
  const g = load();
  let d = g.saveRecord({ date: '2026-10-31', session: '朝', r1: {}, r2: {}, weight: 70, meds: {} });
  assert.strictEqual(d.records.length, 0);
  assert.strictEqual(d.daily[0].weight, 70);
  assert.throws(() => g.saveRecord({ date: '2026-10-30', session: '朝', r1: {}, weight: '', meds: { bp: null } }), /保存する内容/);
  assert.throws(() => g.saveRecord({ date: '2026-10-30', session: '朝', weight: '650.5' }), /範囲外/);
  assert.throws(() => g.saveRecord({ date: '2026-10-30', session: '朝', weight: '65.123' }), /小数点以下2桁/);
});

test('登録を変えると日次シートの薬品名・量が追従する', () => {
  const g = load();
  let d = g.saveRecord({ date: '2026-10-20', session: '朝', weight: 66, meds: {} });
  d = g.saveMedPeriod(period('lipid', 'ロスバスタチン', '2.5mg', '2026-10-01'));
  const rows = g.__mockSpreadsheet.getSheetByName('日次').rows;
  assert.deepStrictEqual(plain(rows[1].slice(5, 8)), ['', 'ロスバスタチン', '2.5mg']);
  g.deleteMedPeriod(d.periods[0].id);
  assert.deepStrictEqual(plain(rows[1].slice(5, 8)), ['', '', '']);
});

test('体重・服薬の項目を送らない保存では日次を変更しない', () => {
  const g = load();
  g.saveRecord({ date: '2026-10-31', session: '朝', r1: BP, weight: 64, meds: {} });
  const d = g.saveRecord({ date: '2026-10-31', session: '夜', r1: BP });
  assert.strictEqual(d.daily[0].weight, 64);
});

// ---- 月平均・PDF ---------------------------------------------------------

const PNG = 'data:image/png;base64,AAAA';
const bp = (sys, dia, pulse) => ({ sys, dia, pulse });

test('月平均：各日の2回平均を朝・夜ごとに月で平均。取り込み値が優先される', () => {
  const g = load();
  g.saveRecord({ date: '2026-09-01', session: '朝', r1: bp(130, 80, 70), r2: bp(126, 78, 66) }); // 128 / 79 / 68
  g.saveRecord({ date: '2026-09-02', session: '朝', r1: bp(120, 76, 60) });
  g.saveRecord({ date: '2026-09-02', session: '夜', r1: bp(118, 72, 64) });
  g.saveRecord({ date: '2026-10-01', session: '朝', r1: bp(140, 90, 70) });
  const sheet = g.getMonthlySheet_();
  sheet.getRange(2, 1, 3, 7).setValues([
    ['2015-01', 135, 85, 70, 125, 78, 72],
    [new Date(Date.UTC(2015, 1, 1) - 9 * 3600e3), 134, 84, '', '', '', ''], // 日付型で入っていても読める
    ['2026/10', 150, 95, 75, '', '', '']                                     // 「2026/10」形式。アプリの値より優先
  ]);
  const m = plain(g.getData().monthly);
  assert.deepStrictEqual(m.map((x) => x.month + ':' + x.source), ['2015-01:import', '2015-02:import', '2026-09:app', '2026-10:import']);
  assert.deepStrictEqual(m[2].morning, { sys: 124, dia: 77.5, pulse: 64, days: 2 });
  assert.deepStrictEqual(m[2].night, { sys: 118, dia: 72, pulse: 64, days: 1 });
  assert.strictEqual(m[1].night, null);
  assert.strictEqual(m[3].morning.sys, 150);
});

test('月ごとの PDF：1枚目に全日の表、2枚目に月平均とグラフ', () => {
  const g = load();
  g.saveRecord({ date: '2026-10-03', session: '朝', r1: bp(130, 80, 70), r2: bp(126, 78, 66) });
  g.saveRecord({ date: '2026-10-03', session: '夜', r1: {}, r2: bp(118, 72, 64) });
  g.saveRecord({ date: '2026-09-30', session: '朝', r1: bp(120, 70, 60) });
  const res = g.createMonthlyReportPdf('2026-10', { bp: PNG, pulse: PNG });
  assert.strictEqual(res.name, '血圧記録_2026-10.pdf');
  const doc = g.__docs[0];
  const items = doc.items;
  assert.strictEqual(items[0].p.text, '血圧・脈拍 測定記録　2026年10月');
  const t = items[1].rows;
  assert.strictEqual(t.length, 2 + 31);
  assert.deepStrictEqual(plain(t[0]), ['', '', '', '朝1回目', '', '', '朝2回目', '', '', '夜1回目', '', '', '夜2回目', '']);
  assert.deepStrictEqual(plain(t[2 + 2]), ['3', '土', '130', '80', '70', '126', '78', '66', '', '', '', '118', '72', '64']);
  assert.deepStrictEqual(plain(t[2]), ['1', '木', '', '', '', '', '', '', '', '', '', '', '', '']);
  assert.ok(items.some((x) => x.type === 'break'));
  const avg = items.find((x, i) => x.type === 'table' && i > 1).rows;
  assert.deepStrictEqual(plain(avg), [['', '測定日数', '最高', '最低', '脈拍'], ['朝', '1', '128.0', '79.0', '68.0'], ['夜', '1', '118.0', '72.0', '64.0']]);
  assert.strictEqual(items.filter((x) => x.type === 'p' && x.p.images.length).length, 2);
  assert.ok(doc.log.includes('setPageWidth') && doc.log.includes('setPageHeight'));
  // 一時ドキュメントはゴミ箱へ
  assert.ok(g.__files.filter((f) => f.temp).every((f) => f.trashed));
  assert.throws(() => g.createMonthlyReportPdf('2026-08', { bp: PNG, pulse: PNG }), /記録がありません/);
  assert.throws(() => g.createMonthlyReportPdf('2026-13', {}), /年月/);
});

test('同じ名前の PDF は作り直すと古い方がゴミ箱へ、フォルダは使い回す', () => {
  const g = load();
  g.saveRecord({ date: '2026-10-03', session: '朝', r1: bp(130, 80, 70) });
  g.createMonthlyReportPdf('2026-10', { bp: PNG, pulse: PNG });
  g.createMonthlyReportPdf('2026-10', { bp: PNG, pulse: PNG });
  const pdfs = g.__files.filter((f) => f.name === '血圧記録_2026-10.pdf');
  assert.deepStrictEqual(plain(pdfs.map((f) => f.trashed)), [true, false]);
  assert.strictEqual(g.__files.filter((f) => f.folder).length, 1);
});

test('月平均グラフの PDF：1ページに2段、3段目から次のページ', () => {
  const g = load();
  const panels = ['2015年〜2017年', '2018年〜2020年', '2021年〜2023年', '2024年〜2026年'].map((title) => ({ title, image: PNG }));
  const res = g.createMonthlyChartPdf(panels, '2015-2026');
  assert.strictEqual(res.name, '血圧月平均グラフ_2015-2026.pdf');
  const kinds = g.__docs[0].items.map((x) => (x.type === 'break' ? 'BREAK' : x.p.images.length ? 'IMG' : x.p.text));
  assert.deepStrictEqual(plain(kinds), ['2015年〜2017年', 'IMG', '2018年〜2020年', 'IMG', 'BREAK', '2021年〜2023年', 'IMG', '2024年〜2026年', 'IMG']);
  assert.throws(() => g.createMonthlyChartPdf([], '2015-2026'), /データがありません/);
  assert.throws(() => g.createMonthlyChartPdf([{ title: 'x', image: 'nope' }], '2015-2026'), /画像/);
});

test('グラフ画像は本文幅いっぱいに縦横比を保って貼る', () => {
  const g = load();
  g.createMonthlyChartPdf([{ title: 't', image: PNG }], '2015-2026');
  const img = g.__docs[0].items.find((x) => x.type === 'p' && x.p.images.length).p.images[0];
  assert.strictEqual(img.height, Math.round((595.28 - 72) * 0.5));
});
