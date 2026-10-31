// テスト用：Code.gs が使う Apps Script API の最小限のインメモリ実装（タイムゾーンは JST 固定）
(function (root) {
  var JST = 9 * 3600 * 1000;

  function MockSheet(ss, name) {
    this.ss = ss; this.name = name; this.rows = [];
  }
  MockSheet.prototype.getParent = function () { return this.ss; };
  MockSheet.prototype.getLastRow = function () {
    for (var i = this.rows.length - 1; i >= 0; i--) {
      if (this.rows[i] && this.rows[i].some(function (v) { return v !== '' && v !== undefined; })) return i + 1;
    }
    return 0;
  };
  MockSheet.prototype.setFrozenRows = function () {};
  MockSheet.prototype.deleteRow = function (r) { this.rows.splice(r - 1, 1); };
  MockSheet.prototype.getRange = function (a, b, c, d) {
    if (typeof a === 'string') return new MockRange(this, 1, 1, 0, 0);
    return new MockRange(this, a, b, c || 1, d || 1);
  };

  function MockRange(sheet, r, c, nr, nc) { this.s = sheet; this.r = r; this.c = c; this.nr = nr; this.nc = nc; }
  MockRange.prototype.setValues = function (vals) {
    for (var i = 0; i < this.nr; i++) {
      var row = this.s.rows[this.r - 1 + i] || (this.s.rows[this.r - 1 + i] = []);
      for (var j = 0; j < this.nc; j++) row[this.c - 1 + j] = vals[i][j];
    }
    return this;
  };
  MockRange.prototype.getValues = function () {
    var out = [];
    for (var i = 0; i < this.nr; i++) {
      var row = this.s.rows[this.r - 1 + i] || [];
      var o = [];
      for (var j = 0; j < this.nc; j++) {
        var v = row[this.c - 1 + j];
        o.push(v === undefined || (typeof v === 'string' && v[0] === '=') ? '' : v);
      }
      out.push(o);
    }
    return out;
  };
  MockRange.prototype.setFormulasR1C1 = function (f) { return this.setValues(f); };
  MockRange.prototype.setFontWeight = function () { return this; };
  MockRange.prototype.setNumberFormat = function () { return this; };
  MockRange.prototype.sort = function (specs) {
    var rows = this.s.rows.slice(this.r - 1, this.r - 1 + this.nr);
    var key = function (v) { return v instanceof Date ? v.getTime() : v; };
    rows.sort(function (x, y) {
      for (var k = 0; k < specs.length; k++) {
        var a = key(x[specs[k].column - 1]), b = key(y[specs[k].column - 1]);
        if (a !== b) return (a < b ? -1 : 1) * (specs[k].ascending ? 1 : -1);
      }
      return 0;
    });
    Array.prototype.splice.apply(this.s.rows, [this.r - 1, rows.length].concat(rows));
    return this;
  };

  function MockSpreadsheet() { this.sheets = {}; }
  MockSpreadsheet.prototype.getSheetByName = function (n) { return this.sheets[n] || null; };
  MockSpreadsheet.prototype.insertSheet = function (n) { return (this.sheets[n] = new MockSheet(this, n)); };
  MockSpreadsheet.prototype.getSpreadsheetTimeZone = function () { return 'Asia/Tokyo'; };

  var ss = new MockSpreadsheet();
  root.__mockSpreadsheet = ss;
  root.SpreadsheetApp = { getActiveSpreadsheet: function () { return ss; } };
  root.LockService = { getScriptLock: function () { return { waitLock: function () {}, releaseLock: function () {} }; } };
  root.Utilities = {
    formatDate: function (d, tz, pattern) {
      var j = new Date(d.getTime() + JST);
      var p = function (n) { return ('0' + n).slice(-2); };
      if (pattern === 'yyyy-MM-dd') return j.getUTCFullYear() + '-' + p(j.getUTCMonth() + 1) + '-' + p(j.getUTCDate());
      if (pattern === 'H') return String(j.getUTCHours());
      throw new Error('unsupported pattern ' + pattern);
    },
    parseDate: function (s, tz, pattern) {
      var p = s.split('-').map(Number);
      return new Date(Date.UTC(p[0], p[1] - 1, p[2]) - JST);
    }
  };
})(typeof globalThis !== 'undefined' ? globalThis : this);
