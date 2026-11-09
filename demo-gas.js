/* 試看模式（?demo=1，只在本機）：在瀏覽器裡模擬 Google 試算表，直接執行 gas/Code.gs＋gas/Seed.gs，
 * 所以試看模式的規則和正式版完全一樣。學生作答相關的工作表存在 localStorage。 */
window.DemoGas = (function () {
  var KEEP = ["學生", "作答紀錄", "錯題", "盲點分析", "題目分析"], KEY = "grammarDemo_v3";

  function makeSheet(name) {
    var sh = { name: name, data: [] };
    function lastRow() { for (var i = sh.data.length - 1; i >= 0; i--) if (sh.data[i] && sh.data[i].some(function (v) { return v !== "" && v != null; })) return i + 1; return 0; }
    function lastCol() { return sh.data.reduce(function (m, r) { return Math.max(m, (r || []).length); }, 0); }
    function cell(r, c) { var row = sh.data[r - 1], v = row ? row[c - 1] : undefined; return v == null ? "" : v; }
    function set(r, c, v) { while (sh.data.length < r) sh.data.push([]); var row = sh.data[r - 1]; while (row.length < c) row.push(""); row[c - 1] = v; }
    function range(r, c, nr, nc) {
      var o = {
        getValues: function () { var out = []; for (var i = 0; i < nr; i++) { var row = []; for (var j = 0; j < nc; j++) row.push(cell(r + i, c + j)); out.push(row); } return out; },
        setValues: function (v) { v.forEach(function (row, i) { row.forEach(function (x, j) { set(r + i, c + j, x); }); }); return o; },
        setValue: function (x) { set(r, c, x); return o; },
        clearContent: function () { for (var i = 0; i < nr; i++) for (var j = 0; j < nc; j++) if (sh.data[r + i - 1]) set(r + i, c + j, ""); return o; },
        setNumberFormat: function () { return o; }, setFontWeight: function () { return o; }, setBackground: function () { return o; }
      };
      return o;
    }
    sh.getLastRow = lastRow; sh.getLastColumn = lastCol; sh.setFrozenRows = function () {};
    sh.getRange = function (a, b, c, d) {
      if (typeof a === "string") return range(1, a.charCodeAt(0) - 64, Math.max(lastRow(), 1), 1);
      return range(a, b, c || 1, d || 1);
    };
    sh.getDataRange = function () { return range(1, 1, Math.max(lastRow(), 1), Math.max(lastCol(), 1)); };
    sh.appendRow = function (row) { var r = lastRow() + 1; row.forEach(function (x, j) { set(r, j + 1, x); }); };
    return sh;
  }

  var sheets = {}, G = null, todayFn = null;
  var ss = { getSheetByName: function (n) { return sheets[n] || null; }, insertSheet: function (n) { return (sheets[n] = makeSheet(n)); } };
  var env = {
    SpreadsheetApp: { getActive: function () { return ss; }, flush: function () {} },
    LockService: { getScriptLock: function () { return { waitLock: function () {}, releaseLock: function () {} }; } },
    Utilities: { formatDate: function (d) { return new Date(d.getTime() + 8 * 3600e3).toISOString().slice(0, 10); } },
    ContentService: { createTextOutput: function (s) { return { setMimeType: function () { return s; } }; }, MimeType: { JSON: "json" } }
  };

  function save() {
    var out = {};
    KEEP.forEach(function (n) {
      out[n] = sheets[n].data.map(function (r) { return r.map(function (v) { return v instanceof Date ? { $d: v.toISOString() } : v; }); });
    });
    try { localStorage.setItem(KEY, JSON.stringify(out)); } catch (e) {}
  }
  function restore() {
    var x = null;
    try { x = JSON.parse(localStorage.getItem(KEY)); } catch (e) {}
    if (!x) return;
    KEEP.forEach(function (n) {
      if (x[n]) sheets[n].data = x[n].map(function (r) { return r.map(function (v) { return v && v.$d ? new Date(v.$d) : v; }); });
    });
  }

  return {
    load: function (getToday) {
      todayFn = getToday;
      return Promise.all(["gas/Code.gs", "gas/Seed.gs"].map(function (f) {
        return fetch(f + "?t=" + Date.now()).then(function (r) { if (!r.ok) throw new Error("讀不到 " + f); return r.text(); });
      })).then(function (src) {
        var names = Object.keys(env);
        var body = src.join("\n") + "\nreturn { doPost: doPost, setup: setup, setToday: function (f) { today = f; } };";
        G = Function.apply(null, names.concat(body)).apply(null, names.map(function (k) { return env[k]; }));
        G.setup();
        restore();
        G.setToday(function () { return todayFn(); });
      });
    },
    call: function (action, data) {
      data = data || {};
      data.action = action;
      var res = JSON.parse(G.doPost({ postData: { contents: JSON.stringify(data) } }));
      if (action === "answer") save();
      return res;
    },
    reset: function () {
      try { localStorage.removeItem(KEY); } catch (e) {}
      KEEP.forEach(function (n) { sheets[n].data = []; });
      G.setup();
    }
  };
})();
