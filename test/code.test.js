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
  assert.throws(() => g.saveRecord({ date: '2026-10-31', session: '朝', r1: {}, r2: {} }), /どちらか/);
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
