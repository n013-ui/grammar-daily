/**
 * 英文文法每日練習 — Google Apps Script 後端
 * 綁在 Google 試算表；第一次請在編輯器執行 setup() 建立工作表並匯入題庫（Seed.gs）。
 * 前端（GitHub Pages）以 POST text/plain JSON 呼叫本網頁應用程式。
 */

var TZ = 'Asia/Taipei';
var SH = {
  SET: '設定', UNIT: '文法說明', Q: '題目', VAR: '變化題', SCHED: '進度',
  STU: '學生', LOG: '作答紀錄', WRONG: '錯題', KP: '考點', BLIND: '盲點分析', QSTAT: '題目分析'
};
var HEAD = {};
HEAD[SH.SET]   = ['項目', '值', '說明'];
HEAD[SH.UNIT]  = ['單元ID', '課次', '標題', '說明HTML'];
HEAD[SH.Q]     = ['題目ID', '課次', '單元', '出處', '題目', 'A', 'B', 'C', 'D', '答案', '解析', '考點', '選項盲點', '文章'];
HEAD[SH.VAR]   = ['原題ID', '序號', '題目', 'A', 'B', 'C', 'D', '答案', '解析'];
HEAD[SH.SCHED] = ['日期', '標題', '文法單元', '題目ID'];
HEAD[SH.STU]   = ['登入名稱', '顯示名稱', '班級', '座號', '總點數', '完成天數', '錯題待訂正', '最後作答'];
HEAD[SH.LOG]   = ['時間', '學號', '姓名', '模式', '練習日期', '題目ID', '選答', '對錯', '得分', '選項內容', '正確答案', '考點', '盲點'];
HEAD[SH.WRONG] = ['學號', '題目ID', '錯誤次數', '狀態', '加入時間', '最後更新', '回收練習次數', '回收紀錄'];
HEAD[SH.KP]    = ['考點ID', '課次', '單元', '考點', '盲點提醒'];
HEAD[SH.BLIND] = ['學號', '姓名', '考點ID', '考點', '作答數', '答錯數', '錯誤率', '盲點統計', '最近答錯（最新在前）'];
HEAD[SH.QSTAT] = ['題目ID', '考點', '題目', '正確答案', '作答數', '答對數', '答對率', '選A', '選B', '選C', '選D'];

// ================= 安裝／升級 =================
function setup() {
  var ss = SpreadsheetApp.getActive();
  setupSheets();
  var S = SEED;
  var set = ss.getSheetByName(SH.SET);
  if (set.getLastRow() < 2) {
    set.getRange(2, 1, 2, 3).setValues([
      ['標題', S.settings.title, '網頁上方顯示的名稱'],
      ['錯題回收題數', S.settings.reviewCount, '每次錯題回收最多幾題']
    ]);
  }
  var u = ss.getSheetByName(SH.UNIT);
  if (u.getLastRow() < 2) u.getRange(2, 1, S.units.length, 4).setValues(S.units.map(function (x) { return [x.id, x.lesson, x.title, x.html.trim()]; }));
  var kp = ss.getSheetByName(SH.KP);
  if (kp.getLastRow() < 2) kp.getRange(2, 1, S.kps.length, 5).setValues(S.kps.map(function (k) { return [k.id, k.lesson, k.unit, k.name, k.tip]; }));
  addMissingQuestions();
  addMissingVariants();
  if (ss.getSheetByName(SH.SCHED).getLastRow() < 2) resetSchedule();
  var st = ss.getSheetByName(SH.STU);
  if (st.getLastRow() < 2) st.getRange(2, 1, 1, 4).setValues([['queena', 'Queena', '', '']]);
}
function setupSheets() {
  var ss = SpreadsheetApp.getActive();
  Object.keys(HEAD).forEach(function (name) {
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    var h = HEAD[name];
    sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground('#e3f1ef');
    sh.setFrozenRows(1);
  });
  var stu = ss.getSheetByName(SH.STU);
  stu.getRange('A:A').setNumberFormat('@');
  stu.getRange(1, 1).setNote('【新增帳號】在下面空白列的 A 欄輸入登入名稱（例如 queena），就能用這個名稱登入，不需要密碼，大小寫不拘。\n' +
    'B 欄「顯示名稱」是網頁上看到的名字，可以不填。\n' +
    'E～H 欄（點數、完成天數…）由系統自動填，請不要手動修改。\n' +
    '刪掉一列＝取消這個帳號（作答紀錄會留著）。');
  stu.getRange(1, 5, 1, 4).setBackground('#eeeeee');
  ss.getSheetByName(SH.SCHED).getRange('A:A').setNumberFormat('@');
  ss.getSheetByName(SH.LOG).getRange('B:B').setNumberFormat('@');
  ss.getSheetByName(SH.LOG).getRange('E:E').setNumberFormat('@');
  ss.getSheetByName(SH.WRONG).getRange('A:A').setNumberFormat('@');
  ss.getSheetByName(SH.BLIND).getRange('A:A').setNumberFormat('@');
  ss.getSheetByName(SH.BLIND).getRange('G:G').setNumberFormat('0%');
  ss.getSheetByName(SH.QSTAT).getRange('G:G').setNumberFormat('0%');
}

// 把 Seed 裡有、試算表裡還沒有的題目補進去（不會覆蓋老師改過的題目）
function addMissingQuestions() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SH.Q), have = {};
  rows(SH.Q).forEach(function (r) { have[str(r[0])] = 1; });
  var add = SEED.questions.filter(function (x) { return !have[x.id]; }).map(function (x) {
    return qRow(x);
  });
  if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, HEAD[SH.Q].length).setValues(add);
  return add.length;
}
function qRow(x) {
  return [x.id, x.lesson, x.unit, x.src, x.q, x.opts[0], x.opts[1], x.opts[2], x.opts[3], x.ans, x.exp,
    x.kp || '', x.traps ? JSON.stringify(x.traps) : '', x.passage || ''];
}
// 老師用：用 Seed 覆蓋「題目、變化題、文法說明、考點、進度」（學生、作答紀錄、錯題不動）
function reloadFromSeed() {
  var ss = SpreadsheetApp.getActive(), S = SEED;
  setupSheets();
  function overwrite(name, data) {
    var sh = ss.getSheetByName(name), w = HEAD[name].length;
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, Math.max(w, sh.getLastColumn())).clearContent();
    if (data.length) sh.getRange(2, 1, data.length, w).setValues(data);
  }
  overwrite(SH.Q, S.questions.map(qRow));
  var vr = [];
  Object.keys(S.variants || {}).forEach(function (id) {
    S.variants[id].forEach(function (v, i) { vr.push([id, i + 1, v[0], v[1][0], v[1][1], v[1][2], v[1][3], v[2], v[3]]); });
  });
  overwrite(SH.VAR, vr);
  overwrite(SH.UNIT, S.units.map(function (x) { return [x.id, x.lesson, x.title, x.html.trim()]; }));
  overwrite(SH.KP, S.kps.map(function (k) { return [k.id, k.lesson, k.unit, k.name, k.tip]; }));
  resetSchedule();
  var set = ss.getSheetByName(SH.SET), sv = set.getDataRange().getValues();
  for (var i = 1; i < sv.length; i++) if (str(sv[i][0]) === '標題') set.getRange(i + 1, 2).setValue(S.settings.title);
  return { questions: S.questions.length, variants: vr.length, days: S.schedule.length };
}
function addMissingVariants() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SH.VAR), have = {};
  rows(SH.VAR).forEach(function (r) { have[str(r[0]) + '#' + str(r[1])] = 1; });
  var add = [];
  Object.keys(SEED.variants || {}).forEach(function (id) {
    SEED.variants[id].forEach(function (v, i) {
      if (!have[id + '#' + (i + 1)]) add.push([id, i + 1, v[0], v[1][0], v[1][1], v[1][2], v[1][3], v[2], v[3]]);
    });
  });
  if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, 9).setValues(add);
  return add.length;
}

// 老師用：用 Seed.gs 的進度覆蓋「進度」工作表（作答紀錄不受影響）
function resetSchedule() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SH.SCHED);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 4).clearContent();
  sh.getRange('A:A').setNumberFormat('@');
  sh.getRange(2, 1, SEED.schedule.length, 4).setValues(SEED.schedule.map(function (x) {
    return [x.date, x.title, x.units.join(','), x.qids.join(',')];
  }));
}

// ================= 入口 =================
function doGet() {
  return json({ ok: true, msg: '英文文法每日練習 API 運作中' });
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
  login: function (p) { return summary(auth(p)); },
  summary: function (p) { return summary(auth(p)); },

  // 每一天的題目：可以提前做、也可以補做（點數一樣算）
  getDay: function (p) {
    var s = auth(p), day = findDay(p.date);
    if (!day) throw new Error('這一天沒有排練習');
    var Q = questionMap(), K = kpMap(), recs = studentLog(s.sid), a = (recs.daily[day.date] || {}), answered = {};
    day.qids.forEach(function (id) {
      if (a[id] && Q[id]) answered[id] = feedback(Q[id], K, a[id].choice, a[id].correct);
    });
    return {
      ok: true, date: day.date, title: day.title, units: unitsFor(day, Q),
      questions: day.qids.filter(function (id) { return Q[id]; }).map(function (id) { return pub(Q[id], K); }),
      answered: answered
    };
  },

  // 錯題回收：任何一天都能練，但只出「前一天以前」答錯的題目；換句話考同一個考點，選項順序打亂
  getReview: function (p) {
    var s = auth(p), Q = questionMap(), K = kpMap(), V = variantMap(), t = today();
    var items = wrongItems(s.sid).filter(function (w) { return w.status === '待訂正' && w.lastDay < t && Q[w.qid]; })
      .slice(0, settings().reviewCount);
    var qs = items.map(function (w) {
      var src = pickVersion(Q[w.qid], V[w.qid], w.attempts, s.sid, Q);
      var order = shuffledIdx(s.sid + '|' + w.qid + '|' + w.attempts);
      var k = K[Q[w.qid].kp] || {};
      return { id: w.qid, vk: src.vk, q: src.q, passage: src.passage || '', opts: order.map(function (i) { return src.opts[i]; }),
        unit: Q[w.qid].unit, kpName: k.name || '', src: src.vk === 'o' ? '原題' : '變化題',
        wrongCount: w.count, attempts: w.attempts, hist: w.hist };
    });
    return { ok: true, title: '錯題回收區', units: unitsFor({ units: [], qids: items.map(function (w) { return w.qid; }) }, Q),
      questions: qs, answered: {} };
  },

  // 錯題本：列出所有錯過的題目與回收紀錄
  wrongBook: function (p) {
    var s = auth(p), Q = questionMap(), K = kpMap(), t = today();
    var list = wrongItems(s.sid).filter(function (w) { return Q[w.qid]; }).map(function (w) {
      var q = Q[w.qid], kn = (K[q.kp] || {}).name || '';
      return { qid: w.qid, q: q.passage ? '克漏字（' + kn + '）' : q.q, kpName: kn,
        count: w.count, status: w.status, added: w.addDay, attempts: w.attempts, hist: w.hist,
        ready: w.status === '待訂正' && w.lastDay < t };
    });
    list.sort(function (a, b) {
      var pa = a.status === '待訂正' ? 0 : 1, pb = b.status === '待訂正' ? 0 : 1;
      return pa - pb || (a.added < b.added ? -1 : a.added > b.added ? 1 : 0);
    });
    return { ok: true, items: list };
  },

  answer: function (p) {
    var s = auth(p), Q = questionMap(), K = kpMap(), q = Q[p.qid];
    if (!q) throw new Error('找不到題目');
    var mode = p.mode === 'review' ? 'review' : p.mode === 'practice' ? 'practice' : 'daily';
    var lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      var ss = SpreadsheetApp.getActive(), now = new Date(), t = today(), gained = 0, correct, choice;
      var wsh = ss.getSheetByName(SH.WRONG), wv = wsh.getDataRange().getValues(), wrow = -1;
      for (var i = 1; i < wv.length; i++) {
        if (str(wv[i][0]) === s.sid && str(wv[i][1]) === p.qid && str(wv[i][3]) === '待訂正') { wrow = i + 1; break; }
      }

      if (mode === 'review') {
        if (wrow < 0) throw new Error('這題已經不在錯題本裡了');
        var w = wv[wrow - 1];
        if (str(w[5]) >= t) throw new Error('這題今天剛錯過，明天才能練');
        var src = versionByKey(q, variantMap()[p.qid], Q, p.vk);
        var right = src.opts['ABCD'.indexOf(src.ans)];
        choice = String(p.choiceText || '');
        correct = choice === right;
        var ci = src.opts.indexOf(choice), letter = ci >= 0 ? 'ABCD'.charAt(ci) : '';
        var trap = correct ? '' : trapOf(src.traps, letter);
        var attempts = (Number(w[6]) || 0) + 1, hist = str(w[7]) + (correct ? '✓' : '✗'), count = Number(w[2]) || 1;
        if (correct) { gained = 1; wsh.getRange(wrow, 4).setValue('已訂正'); }
        else { count++; wsh.getRange(wrow, 3).setValue(count); }
        wsh.getRange(wrow, 6).setValue(now);
        wsh.getRange(wrow, 7, 1, 2).setValues([[attempts, hist]]);
        var vlabel = src.vk === 'o' ? '' : src.vk.indexOf('v:') === 0 ? '#' + (Number(src.vk.slice(2)) + 1) : '→' + src.vk.slice(2);
        ss.getSheetByName(SH.LOG).appendRow([now, s.sid, s.name, 'review', '', p.qid + vlabel, letter, correct ? '對' : '錯', gained,
          choice, right, q.kp, trap]);
        recordAnalysis(s, q, null, correct, trap, p.qid + vlabel, choice);
        SpreadsheetApp.flush();
        var sum = refreshStudent(s), k1 = K[q.kp] || {};
        return { ok: true, correct: correct, ansText: right, exp: src.exp, gained: gained, points: sum.points, wrongCount: count, hist: hist,
          kpName: k1.name || '', kpTip: k1.tip || '', trap: trap };
      }

      choice = String(p.choice || '').toUpperCase();
      if ('ABCD'.indexOf(choice) < 0 || choice.length !== 1) throw new Error('選項不正確');
      correct = choice === q.ans;
      if (mode === 'daily') {
        var day = findDay(p.date);
        if (!day || day.qids.indexOf(p.qid) < 0) throw new Error('題目不在這一天的練習裡');
        var prev = (studentLog(s.sid).daily[day.date] || {})[p.qid];
        if (prev) {
          var fb0 = feedback(q, K, prev.choice, prev.correct);
          fb0.ok = true; fb0.gained = 0; fb0.points = points(s.sid); fb0.repeat = true;
          return fb0;
        }
        gained = correct ? 1 : 0;
      }
      if (!correct) {
        if (wrow > 0) { wsh.getRange(wrow, 3).setValue((Number(wv[wrow - 1][2]) || 1) + 1); wsh.getRange(wrow, 6).setValue(now); }
        else wsh.appendRow([s.sid, p.qid, 1, '待訂正', now, now, 0, '']);
      }
      var fb = feedback(q, K, choice, correct), ctext = q.opts['ABCD'.indexOf(choice)];
      ss.getSheetByName(SH.LOG).appendRow([now, s.sid, s.name, mode, String(p.date || ''), p.qid, choice, correct ? '對' : '錯', gained,
        ctext, q.opts['ABCD'.indexOf(q.ans)], q.kp, fb.trap]);
      recordAnalysis(s, q, choice, correct, fb.trap, p.qid, ctext);
      SpreadsheetApp.flush();
      var sum2 = refreshStudent(s);
      fb.ok = true; fb.gained = gained; fb.points = sum2.points;
      return fb;
    } finally {
      lock.releaseLock();
    }
  }
};

// 答題後給學生看的回饋：答案、詳解、考點、盲點提醒、「你可能是…」
function feedback(q, K, choice, correct) {
  var k = K[q.kp] || {};
  return { choice: choice, correct: correct, ans: q.ans, exp: q.exp, kpName: k.name || '', kpTip: k.tip || '',
    trap: correct ? '' : trapOf(q.traps, choice) };
}
function trapOf(traps, letter) {
  if (!letter) return '';
  return (traps && traps[letter]) || '';
}

// ================= 老師分析：盲點分析、題目分析 =================
function recordAnalysis(s, q, letter, correct, trap, qlabel, choiceText) {
  var ss = SpreadsheetApp.getActive(), K = kpMap(), k = K[q.kp] || {};
  if (q.kp) {
    var b = ss.getSheetByName(SH.BLIND), bv = b.getDataRange().getValues(), r = -1;
    for (var i = 1; i < bv.length; i++) if (str(bv[i][0]) === s.sid && str(bv[i][2]) === q.kp) { r = i; break; }
    var row = r >= 0 ? bv[r] : [s.sid, s.name, q.kp, k.name || q.kp, 0, 0, 0, '', ''];
    row = blindRow(row, correct, trap, qlabel, choiceText);
    if (r >= 0) b.getRange(r + 1, 1, 1, row.length).setValues([row]); else b.appendRow(row);
  }
  if (letter) {
    var qs = ss.getSheetByName(SH.QSTAT), qv = qs.getDataRange().getValues(), r2 = -1;
    for (var j = 1; j < qv.length; j++) if (str(qv[j][0]) === q.id) { r2 = j; break; }
    var qr = r2 >= 0 ? qv[r2] : [q.id, k.name || q.kp, (q.passage ? '［克漏字］' : '') + q.q.slice(0, 60), q.ans, 0, 0, 0, 0, 0, 0, 0];
    qr = qstatRow(qr, letter, correct);
    if (r2 >= 0) qs.getRange(r2 + 1, 1, 1, qr.length).setValues([qr]); else qs.appendRow(qr);
  }
}
function blindRow(row, correct, trap, qlabel, choiceText) {
  row = row.slice(0, 9);
  row[4] = (Number(row[4]) || 0) + 1;
  if (!correct) {
    row[5] = (Number(row[5]) || 0) + 1;
    var label = trap || '其他（未分類）', cnt = {}, order = [];
    str(row[7]).split('；').forEach(function (x) {
      var m = x.match(/^(.*)×(\d+)$/);
      if (m) { cnt[m[1]] = Number(m[2]); order.push(m[1]); }
    });
    if (!cnt[label]) order.push(label);
    cnt[label] = (cnt[label] || 0) + 1;
    order.sort(function (a, b) { return cnt[b] - cnt[a]; });
    row[7] = order.map(function (x) { return x + '×' + cnt[x]; }).join('；');
    var recent = str(row[8]) ? str(row[8]).split('\n') : [];
    recent.unshift(qlabel + ' 選「' + choiceText + '」' + (trap ? '→' + trap : ''));
    row[8] = recent.slice(0, 5).join('\n');
  }
  row[6] = row[4] ? row[5] / row[4] : 0;
  return row;
}
function qstatRow(qr, letter, correct) {
  qr = qr.slice(0, 11);
  qr[4] = (Number(qr[4]) || 0) + 1;
  if (correct) qr[5] = (Number(qr[5]) || 0) + 1;
  qr[6] = qr[4] ? qr[5] / qr[4] : 0;
  var c = 7 + 'ABCD'.indexOf(letter);
  qr[c] = (Number(qr[c]) || 0) + 1;
  return qr;
}
// 老師用：依「作答紀錄」整個重算兩張分析表
function rebuildAnalysis() {
  var ss = SpreadsheetApp.getActive(), Q = questionMap(), K = kpMap(), B = {}, Bo = [], S = {}, So = [];
  rows(SH.LOG).forEach(function (r) {
    var sid = str(r[1]), qlabel = str(r[5]), qid = qlabel.split(/[#→]/)[0], q = Q[qid];
    if (!sid || !q) return;
    var correct = str(r[7]) === '對', trap = str(r[12]), choiceText = str(r[9]) || str(r[6]), mode = str(r[3]);
    if (q.kp) {
      var key = sid + '|' + q.kp;
      if (!B[key]) { B[key] = [sid, str(r[2]), q.kp, (K[q.kp] || {}).name || q.kp, 0, 0, 0, '', '']; Bo.push(key); }
      B[key] = blindRow(B[key], correct, trap, qlabel, choiceText);
    }
    var letter = str(r[6]);
    if (mode !== 'review' && letter.length === 1 && 'ABCD'.indexOf(letter) >= 0) {
      if (!S[qid]) { S[qid] = [qid, (K[q.kp] || {}).name || q.kp, (q.passage ? '［克漏字］' : '') + q.q.slice(0, 60), q.ans, 0, 0, 0, 0, 0, 0, 0]; So.push(qid); }
      S[qid] = qstatRow(S[qid], letter, correct);
    }
  });
  function write(name, data) {
    var sh = ss.getSheetByName(name), w = HEAD[name].length;
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, w).clearContent();
    if (data.length) sh.getRange(2, 1, data.length, w).setValues(data);
  }
  write(SH.BLIND, Bo.sort().map(function (k) { return B[k]; }));
  write(SH.QSTAT, So.sort().map(function (k) { return S[k]; }));
  return { blind: Bo.length, qstat: So.length };
}

// ================= 資料存取 =================
function rows(name) {
  var sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) return [];
  var v = sh.getDataRange().getValues();
  v.shift();
  return v;
}
function str(x) {
  if (x instanceof Date) return Utilities.formatDate(x, TZ, 'yyyy-MM-dd');
  return String(x == null ? '' : x).trim();
}
function list(x) { return str(x).split(/[,，、\s]+/).filter(String); }
function today() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }

function settings() {
  var m = {};
  rows(SH.SET).forEach(function (r) { m[str(r[0])] = r[1]; });
  return { title: str(m['標題'] || '英文文法每日練習'), reviewCount: Number(m['錯題回收題數']) || 10 };
}

// 只用一個登入名稱（不分大小寫），對「學生」表的學號或姓名欄都可以；不需要密碼
function auth(p) {
  var key = str(p.sid || p.name).toLowerCase();
  if (!key) throw new Error('請輸入登入名稱');
  var r = rows(SH.STU).filter(function (x) { return str(x[0]).toLowerCase() === key || str(x[1]).toLowerCase() === key; })[0];
  if (!r) throw new Error('找不到這個登入名稱');
  return { sid: str(r[0]), name: str(r[1]) || str(r[0]), cls: str(r[2]) };
}

function questionMap() {
  var m = {};
  rows(SH.Q).forEach(function (r) {
    var id = str(r[0]);
    if (!id) return;
    m[id] = { id: id, lesson: str(r[1]), unit: str(r[2]), src: str(r[3]), q: str(r[4]),
      opts: [str(r[5]), str(r[6]), str(r[7]), str(r[8])], ans: str(r[9]).toUpperCase(), exp: str(r[10]),
      kp: str(r[11]), traps: parseJson(r[12]), passage: str(r[13]) };
  });
  return m;
}
function parseJson(x) { try { return x ? JSON.parse(str(x)) : null; } catch (e) { return null; } }
function kpMap() {
  var m = {};
  rows(SH.KP).forEach(function (r) { if (str(r[0])) m[str(r[0])] = { id: str(r[0]), lesson: str(r[1]), unit: str(r[2]), name: str(r[3]), tip: str(r[4]) }; });
  return m;
}
function variantMap() {
  var m = {};
  rows(SH.VAR).filter(function (r) { return str(r[0]) && str(r[2]); })
    .sort(function (a, b) { return (Number(a[1]) || 0) - (Number(b[1]) || 0); })
    .forEach(function (r) {
      var id = str(r[0]);
      (m[id] = m[id] || []).push({ q: str(r[2]), opts: [str(r[3]), str(r[4]), str(r[5]), str(r[6])], ans: str(r[7]).toUpperCase(), exp: str(r[8]) });
    });
  return m;
}
// 第 k 次回收練習用哪個版本（vk）：
//   偶數次（0、2、4…）出「換句話」：有變化題就輪流用變化題，沒有就挑同一個考點的另一題（依學號固定挑）
//   奇數次出原題（選項另外打亂）；連續兩次不會同一句
function pickVersion(q, vars, k, sid, Q) {
  vars = vars || [];
  if (k % 2 === 1) return versionByKey(q, vars, Q, 'o');
  var half = Math.floor(k / 2);
  if (vars.length) return versionByKey(q, vars, Q, 'v:' + (half % vars.length));
  var sib = siblings(q, Q);
  if (!sib.length) return versionByKey(q, vars, Q, 'o');
  var h = shuffledIdx(sid + '#' + q.id)[0] * 7 + q.id.length;
  return versionByKey(q, vars, Q, 'q:' + sib[(h + half) % sib.length].id);
}
// 同考點的其他題目；優先用「沒有排進每日進度」的題目，學生之後才不會在每日練習又遇到
var _schedIds = null;
function siblings(q, Q) {
  if (!_schedIds) { _schedIds = {}; schedule().forEach(function (d) { d.qids.forEach(function (id) { _schedIds[id] = 1; }); }); }
  var all = Object.keys(Q).map(function (id) { return Q[id]; })
    .filter(function (x) { return x.kp && x.kp === q.kp && x.id !== q.id && !x.passage; })
    .sort(function (a, b) { return a.id < b.id ? -1 : 1; });
  var spare = all.filter(function (x) { return !_schedIds[x.id]; });
  return spare.length ? spare : all;
}
function versionByKey(q, vars, Q, vk) {
  vk = String(vk == null || vk === '' ? 'o' : vk);
  var m = vk.match(/^v:(\d+)$/);
  if (m && vars && vars[Number(m[1])]) {
    var v = vars[Number(m[1])];
    return { vk: vk, q: v.q, opts: v.opts, ans: v.ans, exp: v.exp, traps: null };
  }
  m = vk.match(/^q:(.+)$/);
  if (m && Q[m[1]] && Q[m[1]].kp === q.kp) {
    var o = Q[m[1]];
    return { vk: vk, q: o.q, opts: o.opts, ans: o.ans, exp: o.exp, traps: o.traps };
  }
  return { vk: 'o', q: q.q, opts: q.opts, ans: q.ans, exp: q.exp, traps: q.traps, passage: q.passage };
}
function shuffledIdx(seedStr) {
  var h = 2166136261;
  for (var i = 0; i < seedStr.length; i++) { h ^= seedStr.charCodeAt(i); h = Math.imul(h, 16777619) >>> 0; }
  var a = [0, 1, 2, 3];
  for (var j = 3; j > 0; j--) { h = (Math.imul(h, 1103515245) + 12345) >>> 0; var r = h % (j + 1); var t = a[j]; a[j] = a[r]; a[r] = t; }
  return a;
}
function pub(q, K) { return { id: q.id, q: q.q, opts: q.opts, src: q.src, unit: q.unit, passage: q.passage || '', kpName: ((K || {})[q.kp] || {}).name || '' }; }

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

function wrongItems(sid) {
  return rows(SH.WRONG).filter(function (r) { return str(r[0]) === sid && str(r[1]); }).map(function (r) {
    return { qid: str(r[1]), count: Number(r[2]) || 1, status: str(r[3]), addDay: str(r[4]), lastDay: str(r[5]),
      attempts: Number(r[6]) || 0, hist: str(r[7]), added: r[4] };
  }).sort(function (a, b) { return new Date(a.added) - new Date(b.added); });
}

function summary(s) {
  var t = today(), cfg = settings(), log = studentLog(s.sid), W = wrongItems(s.sid);
  var days = schedule().map(function (d) {
    var a = log.daily[d.date] || {}, n = 0, sc = 0;
    d.qids.forEach(function (id) { if (a[id]) { n++; if (a[id].correct) sc++; } });
    var st = n >= d.qids.length ? 'done' : n > 0 ? 'partial' : d.date > t ? 'ahead' : 'todo';
    return { date: d.date, title: d.title, total: d.qids.length, answered: n, score: sc, status: st };
  });
  var pending = W.filter(function (w) { return w.status === '待訂正'; });
  return {
    ok: true, today: t, points: log.points, days: days,
    wrongCount: pending.length,
    reviewReady: pending.filter(function (w) { return w.lastDay < t; }).length,
    reviewWaiting: pending.filter(function (w) { return w.lastDay >= t; }).length,
    fixedCount: W.filter(function (w) { return w.status === '已訂正'; }).length,
    reviewCount: cfg.reviewCount,
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

// 老師用：重新計算所有學生的點數（手動執行）
function refreshAllStudents() {
  rows(SH.STU).forEach(function (r) { if (str(r[0])) refreshStudent({ sid: str(r[0]), name: str(r[1]) }); });
}
