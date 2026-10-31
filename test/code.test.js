// Code.gs のロジックをモック上で検証する： node --test test/
const test = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

function load() {
  const ctx = vm.createContext({ console, Date, Math, Number, String, isFinite, Error, Array, Object });
  vm.runInContext(fs.readFileSync(path.join(__dirname, 'mock-gas.js'), 'utf8'), ctx);
  vm.runInContext(fs.readFileSync(path.join(__dirname, '../src/Code.gs'), 'utf8'), ctx);
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
