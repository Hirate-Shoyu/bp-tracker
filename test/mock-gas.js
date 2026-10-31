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
  MockRange.prototype.clearContent = function () {
    for (var i = 0; i < this.nr; i++) {
      var row = this.s.rows[this.r - 1 + i];
      if (row) for (var j = 0; j < this.nc; j++) row[this.c - 1 + j] = '';
    }
    return this;
  };
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
  var uuid = 0;
  root.Utilities = {
    getUuid: function () { return 'id-' + (++uuid); },
    formatDate: function (d, tz, pattern) {
      var j = new Date(d.getTime() + JST);
      var p = function (n) { return ('0' + n).slice(-2); };
      if (pattern === 'yyyy-MM-dd') return j.getUTCFullYear() + '-' + p(j.getUTCMonth() + 1) + '-' + p(j.getUTCDate());
      if (pattern === 'yyyy-MM') return j.getUTCFullYear() + '-' + p(j.getUTCMonth() + 1);
      if (pattern === 'H') return String(j.getUTCHours());
      throw new Error('unsupported pattern ' + pattern);
    },
    parseDate: function (s, tz, pattern) {
      var p = s.split('-').map(Number);
      return new Date(Date.UTC(p[0], p[1] - 1, p[2]) - JST);
    }
  };
  root.Utilities.newBlob = function (bytes, type, name) { return { bytes: bytes, type: type, name: name }; };
  root.Utilities.base64Decode = function (s) { return 'decoded:' + s.length; };

  // ---- DocumentApp / DriveApp / PropertiesService（PDF 作成の呼び出しを記録する簡易版）----
  // 未定義のメソッドは自分自身を返す（メソッドチェーン用）。記録したい操作だけ個別に実装する。
  function chain(obj, log) {
    return new Proxy(obj, {
      get: function (t, k) {
        if (k in t) return t[k];
        if (typeof k !== 'string') return undefined;
        return function () { log.push(k); return this; };
      }
    });
  }
  function makeParagraph(doc, text) {
    var p = { text: text || '', images: [] };
    var proxy = chain({
      setText: function (t) { p.text = t; return proxy; },
      appendInlineImage: function (blob) {
        var img = { blob: blob };
        p.images.push(img);
        return chain({
          getWidth: function () { return 1000; }, getHeight: function () { return 500; }, // 画像の元サイズ（px）
          setHeight: function (h) { img.height = h; return this; }
        }, doc.log);
      },
      editAsText: function () { return chain({}, doc.log); },
      asParagraph: function () { return proxy; },
      info: p
    }, doc.log);
    return proxy;
  }
  function makeTable(doc, rows) {
    var table = chain({
      rows: rows,
      getNumRows: function () { return rows.length; },
      getRow: function (i) {
        return chain({
          getNumCells: function () { return rows[i].length; },
          getCell: function (j) {
            var para = makeParagraph(doc, rows[i][j]);
            return chain({
              getText: function () { return rows[i][j]; },
              getChild: function () { return para; },
              editAsText: function () { return chain({}, doc.log); }
            }, doc.log);
          }
        }, doc.log);
      }
    }, doc.log);
    return table;
  }
  var docs = [];
  root.__docs = docs;
  root.DocumentApp = {
    HorizontalAlignment: { LEFT: 'LEFT', CENTER: 'CENTER' },
    VerticalAlignment: { CENTER: 'CENTER' },
    create: function (name) {
      var doc = { name: name, id: 'doc' + docs.length, items: [], log: [] };
      docs.push(doc);
      var first = makeParagraph(doc, '');
      doc.items.push({ type: 'p', p: first.info });
      var body = chain({
        getChild: function () { return first; },
        appendParagraph: function (t) { var p = makeParagraph(doc, t); doc.items.push({ type: 'p', p: p.info }); return p; },
        appendTable: function (rows) { doc.items.push({ type: 'table', rows: rows }); return makeTable(doc, rows); },
        appendPageBreak: function () { doc.items.push({ type: 'break' }); return this; }
      }, doc.log);
      return chain({ getId: function () { return doc.id; }, getBody: function () { return body; }, saveAndClose: function () { doc.closed = true; } }, doc.log);
    }
  };
  var files = [];
  root.__files = files;
  function makeFile(name, extra) {
    var f = { name: name, trashed: false };
    for (var k in extra) f[k] = extra[k];
    files.push(f);
    return {
      info: f,
      getName: function () { return f.name; },
      getUrl: function () { return 'https://drive.example/' + encodeURIComponent(f.name); },
      getId: function () { return f.id || f.name; },
      setTrashed: function (v) { f.trashed = v; return this; },
      isTrashed: function () { return f.trashed; },
      getAs: function () { return { setName: function (n) { this.name = n; return this; } }; }
    };
  }
  var folders = {};
  function makeFolder(name) {
    var id = 'folder' + Object.keys(folders).length;
    var folder = makeFile(name, { id: id, folder: true });
    folder.createFile = function (blob) { return makeFile(blob.name, { parent: id }); };
    folder.getFilesByName = function (n) {
      var list = files.filter(function (f) { return f.parent === id && f.name === n && !f.trashed; });
      var i = 0;
      return { hasNext: function () { return i < list.length; }, next: function () { var f = list[i++]; return { setTrashed: function (v) { f.trashed = v; } }; } };
    };
    folders[id] = folder;
    return folder;
  }
  root.DriveApp = {
    getFileById: function (id) { return makeFile('temp:' + id, { temp: true }); },
    createFolder: makeFolder,
    getFolderById: function (id) { if (!folders[id]) throw new Error('not found'); return folders[id]; }
  };
  var props = {};
  root.PropertiesService = { getScriptProperties: function () {
    return { getProperty: function (k) { return props[k] || null; }, setProperty: function (k, v) { props[k] = v; } };
  } };
})(typeof globalThis !== 'undefined' ? globalThis : this);
