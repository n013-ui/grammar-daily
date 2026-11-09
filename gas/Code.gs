/**
 * 英文文法每日十題 — Google Apps Script 後端
 * 綁在 Google 試算表；第一次請在編輯器執行 setup() 建立工作表並匯入題庫（Seed.gs）。
 * 前端（GitHub Pages）以 POST text/plain JSON 呼叫本網頁應用程式。
 */

var TZ = 'Asia/Taipei';
var SH = {
  SET: '設定', UNIT: '文法說明', Q: '題目', SCHED: '進度',
  STU: '學生', LOG: '作答紀錄', WRONG: '錯題'
};
var HEAD = {};
HEAD[SH.SET]   = ['項目', '值', '說明'];
HEAD[SH.UNIT]  = ['單元ID', '課次', '標題', '說明HTML'];
HEAD[SH.Q]     = ['題目ID', '課次', '單元', '出處', '題目', 'A', 'B', 'C', 'D', '答案', '解析'];
HEAD[SH.SCHED] = ['日期', '標題', '文法單元', '題目ID'];
HEAD[SH.STU]   = ['學號', '姓名', '班級', '座號', '總點數', '完成天數', '錯題待訂正', '最後作答'];
HEAD[SH.LOG]   = ['時間', '學號', '姓名', '模式', '練習日期', '題目ID', '選答', '對錯', '得分'];
HEAD[SH.WRONG] = ['學號', '題目ID', '錯誤次數', '狀態', '加入時間', '最後更新'];

// ================= 安裝 =================
function setup() {
  var ss = SpreadsheetApp.getActive();
  Object.keys(HEAD).forEach(function (name) {
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, HEAD[name].length).setValues([HEAD[name]]).setFontWeight('bold').setBackground('#e3f1ef');
      sh.setFrozenRows(1);
    }
  });
  // 學號、日期、題目ID 欄設成純文字，避免被自動轉成數字或日期
  ss.getSheetByName(SH.STU).getRange('A:A').setNumberFormat('@');
  ss.getSheetByName(SH.SCHED).getRange('A:A').setNumberFormat('@');
  ss.getSheetByName(SH.LOG).getRange('B:B').setNumberFormat('@');
  ss.getSheetByName(SH.LOG).getRange('E:E').setNumberFormat('@');
  ss.getSheetByName(SH.WRONG).getRange('A:A').setNumberFormat('@');

  var S = SEED;
  var set = ss.getSheetByName(SH.SET);
  if (set.getLastRow() < 2) {
    set.getRange(2, 1, 3, 3).setValues([
      ['標題', S.settings.title, '網頁上方顯示的名稱'],
      ['錯題回收星期', S.settings.reviewDays.join(','), '0=週日 6=週六，可填多個，用逗號隔開'],
      ['錯題回收題數', S.settings.reviewCount, '每次錯題回收最多幾題']
    ]);
  }
  var u = ss.getSheetByName(SH.UNIT);
  if (u.getLastRow() < 2) u.getRange(2, 1, S.units.length, 4).setValues(S.units.map(function (x) { return [x.id, x.lesson, x.title, x.html.trim()]; }));
  var q = ss.getSheetByName(SH.Q);
  if (q.getLastRow() < 2) q.getRange(2, 1, S.questions.length, 11).setValues(S.questions.map(function (x) {
    return [x.id, x.lesson, x.unit, x.src, x.q, x.opts[0], x.opts[1], x.opts[2], x.opts[3], x.ans, x.exp];
  }));
  var sc = ss.getSheetByName(SH.SCHED);
  if (sc.getLastRow() < 2) sc.getRange(2, 1, S.schedule.length, 4).setValues(S.schedule.map(function (x) {
    return [x.date, x.title, x.units.join(','), x.qids.join(',')];
  }));
  var st = ss.getSheetByName(SH.STU);
  if (st.getLastRow() < 2) st.getRange(2, 1, 1, 4).setValues([['11001', '試看同學', '801', '1']]);
}

// ================= 入口 =================
function doGet() {
  return json({ ok: true, msg: '英文文法每日十題 API 運作中' });
}

function doPost(e) {
  try {
    var p = JSON.parse(e.postData.contents);
    var fn = API[p.action];
    if (!fn) throw new Error('未知的動作');
    return json(fn(p));
  } catch (err) {
    return json({ ok: false, error: err.message || String(err) });
  }
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// ================= API =================
var API = {
  login: function (p) { var s = auth(p); return summary(s); },
  summary: function (p) { var s = auth(p); return summary(s); },

  getDay: function (p) {
    var s = auth(p), day = findDay(p.date);
    if (!day) throw new Error('這一天沒有排練習');
    if (day.date > today()) throw new Error('這一天還沒開放');
    var Q = questionMap(), recs = studentLog(s.sid), a = (recs.daily[day.date] || {}), answered = {};
    day.qids.forEach(function (id) {
      if (a[id] && Q[id]) answered[id] = { choice: a[id].choice, correct: a[id].correct, ans: Q[id].ans, exp: Q[id].exp };
    });
    return {
      ok: true, date: day.date, title: day.title, units: unitsFor(day, Q),
      questions: day.qids.filter(function (id) { return Q[id]; }).map(function (id) { return pub(Q[id]); }),
      answered: answered
    };
  },

  getReview: function (p) {
    var s = auth(p), cfg = settings();
    if (cfg.reviewDays.indexOf(weekday()) < 0) throw new Error('錯題回收在週末開放');
    var Q = questionMap();
    var ids = pendingWrong(s.sid).filter(function (id) { return Q[id]; }).slice(0, cfg.reviewCount);
    return { ok: true, title: '週末錯題回收', units: unitsFor({ units: [], qids: ids }, Q),
      questions: ids.map(function (id) { return pub(Q[id]); }), answered: {} };
  },

  answer: function (p) {
    var s = auth(p), Q = questionMap(), q = Q[p.qid];
    if (!q) throw new Error('找不到題目');
    var choice = String(p.choice || '').toUpperCase();
    if ('ABCD'.indexOf(choice) < 0 || choice.length !== 1) throw new Error('選項不正確');
    var mode = p.mode === 'review' ? 'review' : p.mode === 'practice' ? 'practice' : 'daily';
    var lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      var correct = choice === q.ans, gained = 0, ss = SpreadsheetApp.getActive();
      if (mode === 'daily') {
        var day = findDay(p.date);
        if (!day || day.qids.indexOf(p.qid) < 0) throw new Error('題目不在這一天的練習裡');
        if (day.date > today()) throw new Error('這一天還沒開放');
        var prev = (studentLog(s.sid).daily[day.date] || {})[p.qid];
        if (prev) {
          return { ok: true, correct: prev.correct, ans: q.ans, exp: q.exp, gained: 0, points: points(s.sid), repeat: true };
        }
        gained = correct ? 1 : 0;
      }
      var wsh = ss.getSheetByName(SH.WRONG), wv = wsh.getDataRange().getValues(), now = new Date(), wrow = -1;
      for (var i = 1; i < wv.length; i++) {
        if (String(wv[i][0]) === s.sid && String(wv[i][1]) === p.qid && wv[i][3] === '待訂正') { wrow = i + 1; break; }
      }
      if (mode === 'review') {
        if (wrow > 0) {
          if (correct) { wsh.getRange(wrow, 4).setValue('已訂正'); wsh.getRange(wrow, 6).setValue(now); gained = 1; }
          else { wsh.getRange(wrow, 3).setValue(Number(wv[wrow - 1][2]) + 1); wsh.getRange(wrow, 6).setValue(now); }
        }
      } else if (!correct) {
        if (wrow > 0) { wsh.getRange(wrow, 3).setValue(Number(wv[wrow - 1][2]) + 1); wsh.getRange(wrow, 6).setValue(now); }
        else wsh.appendRow([s.sid, p.qid, 1, '待訂正', now, now]);
      }
      ss.getSheetByName(SH.LOG).appendRow([now, s.sid, s.name, mode, mode === 'review' ? '' : String(p.date || ''), p.qid, choice, correct ? '對' : '錯', gained]);
      SpreadsheetApp.flush();
      var sum = refreshStudent(s);
      return { ok: true, correct: correct, ans: q.ans, exp: q.exp, gained: gained, points: sum.points };
    } finally {
      lock.releaseLock();
    }
  }
};

// ================= 資料存取 =================
function rows(name) {
  var v = SpreadsheetApp.getActive().getSheetByName(name).getDataRange().getValues();
  v.shift();
  return v;
}
function str(x) {
  if (x instanceof Date) return Utilities.formatDate(x, TZ, 'yyyy-MM-dd');
  return String(x == null ? '' : x).trim();
}
function list(x) { return str(x).split(/[,，、\s]+/).filter(String); }
function today() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function weekday() { return Number(Utilities.formatDate(new Date(), TZ, 'u')) % 7; } // 0=日

function settings() {
  var m = {};
  rows(SH.SET).forEach(function (r) { m[str(r[0])] = r[1]; });
  return {
    title: str(m['標題'] || '英文文法每日十題'),
    reviewDays: list(m['錯題回收星期'] === '' || m['錯題回收星期'] == null ? '0,6' : m['錯題回收星期']).map(Number),
    reviewCount: Number(m['錯題回收題數']) || 10
  };
}

function auth(p) {
  var sid = str(p.sid), name = str(p.name);
  if (!sid || !name) throw new Error('請輸入學號和姓名');
  var r = rows(SH.STU).filter(function (x) { return str(x[0]) === sid; })[0];
  if (!r || str(r[1]) !== name) throw new Error('找不到這個學號，或姓名不符，請找老師確認');
  return { sid: sid, name: str(r[1]), cls: str(r[2]) };
}

function questionMap() {
  var m = {};
  rows(SH.Q).forEach(function (r) {
    var id = str(r[0]);
    if (!id) return;
    m[id] = { id: id, lesson: str(r[1]), unit: str(r[2]), src: str(r[3]), q: str(r[4]),
      opts: [str(r[5]), str(r[6]), str(r[7]), str(r[8])], ans: str(r[9]).toUpperCase(), exp: str(r[10]) };
  });
  return m;
}
function pub(q) { return { id: q.id, q: q.q, opts: q.opts, src: q.src, unit: q.unit }; }

function schedule() {
  return rows(SH.SCHED).map(function (r) {
    return { date: str(r[0]), title: str(r[1]), units: list(r[2]), qids: list(r[3]) };
  }).filter(function (d) { return /^\d{4}-\d{2}-\d{2}$/.test(d.date) && d.qids.length; })
    .sort(function (a, b) { return a.date < b.date ? -1 : 1; });
}
function findDay(date) { return schedule().filter(function (d) { return d.date === str(date); })[0]; }

function unitsFor(day, Q) {
  var ids = day.units.slice();
  if (!ids.length) day.qids.forEach(function (id) { var u = Q[id] && Q[id].unit; if (u && ids.indexOf(u) < 0) ids.push(u); });
  var U = {};
  rows(SH.UNIT).forEach(function (r) { U[str(r[0])] = { id: str(r[0]), lesson: str(r[1]), title: str(r[2]), html: str(r[3]) }; });
  return ids.sort().map(function (id) { return U[id]; }).filter(Boolean);
}

function studentLog(sid) {
  var daily = {}, pts = 0;
  rows(SH.LOG).forEach(function (r) {
    if (str(r[1]) !== sid) return;
    pts += Number(r[8]) || 0;
    if (str(r[3]) === 'daily') {
      var d = str(r[4]), id = str(r[5]);
      daily[d] = daily[d] || {};
      if (!daily[d][id]) daily[d][id] = { choice: str(r[6]), correct: str(r[7]) === '對' };
    }
  });
  return { daily: daily, points: pts };
}
function points(sid) { return studentLog(sid).points; }

function pendingWrong(sid) {
  return rows(SH.WRONG).filter(function (r) { return str(r[0]) === sid && str(r[3]) === '待訂正'; })
    .sort(function (a, b) { return new Date(a[4]) - new Date(b[4]); })
    .map(function (r) { return str(r[1]); });
}

function summary(s) {
  var t = today(), cfg = settings(), log = studentLog(s.sid);
  var days = schedule().map(function (d) {
    var a = log.daily[d.date] || {}, n = 0, sc = 0;
    d.qids.forEach(function (id) { if (a[id]) { n++; if (a[id].correct) sc++; } });
    var st = d.date > t ? 'locked' : n >= d.qids.length ? 'done' : n > 0 ? 'partial' : 'todo';
    return { date: d.date, title: d.title, total: d.qids.length, answered: n, score: sc, status: st };
  });
  return {
    ok: true, today: t, points: log.points, wrongCount: pendingWrong(s.sid).length, days: days,
    reviewOpen: cfg.reviewDays.indexOf(weekday()) >= 0, reviewDays: cfg.reviewDays,
    doneDays: days.filter(function (d) { return d.status === 'done'; }).length,
    student: { sid: s.sid, name: s.name }, title: cfg.title
  };
}

// 更新「學生」表上的點數、完成天數、錯題數，讓老師直接看
function refreshStudent(s) {
  var sum = summary(s);
  var sh = SpreadsheetApp.getActive().getSheetByName(SH.STU), v = sh.getRange('A:A').getValues();
  for (var i = 1; i < v.length; i++) {
    if (str(v[i][0]) === s.sid) {
      sh.getRange(i + 1, 5, 1, 4).setValues([[sum.points, sum.doneDays, sum.wrongCount, new Date()]]);
      break;
    }
  }
  return sum;
}

// 老師用：用 Seed.gs 的進度覆蓋「進度」工作表（手動執行；作答紀錄不受影響）
function resetSchedule() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SH.SCHED);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 4).clearContent();
  sh.getRange('A:A').setNumberFormat('@');
  sh.getRange(2, 1, SEED.schedule.length, 4).setValues(SEED.schedule.map(function (x) {
    return [x.date, x.title, x.units.join(','), x.qids.join(',')];
  }));
}

// 老師用：重新計算所有學生的點數（手動執行）
function refreshAllStudents() {
  rows(SH.STU).forEach(function (r) { if (str(r[0])) refreshStudent({ sid: str(r[0]), name: str(r[1]) }); });
}
