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
  if (!ss.getSheetByName(TOOL)) setupTeacherTool();
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
  if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, HEAD[SH.Q].length).setNumberFormat('@').setValues(add);
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
    // 題目、選項一律存成純文字，避免試算表把「6:10」這類內容自動變成時間或數字
    if (data.length) sh.getRange(2, 1, data.length, w).setNumberFormat('@').setValues(data);
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
  if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, 9).setNumberFormat('@').setValues(add);
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
      if (p.mode === 'drill') return answerDrill(s, q, K, p);
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
    student: { sid: s.sid, name: s.name }, title: cfg.title, drill: drillCounts(s.sid, days)
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

// ================= 考前特訓（學生）＋老師工具（錯題報告、考前練習卷） =================
var LESSON_NAME = { L3: '第三課', L4: '第四課', ALL: '第三、四課' };
function lessonKey(x) {
  x = str(x);
  if (/三|L3|^3$/.test(x)) return 'L3';
  if (/四|L4|^4$/.test(x)) return 'L4';
  return 'ALL';
}
function inLesson(q, lesson) { return lesson === 'ALL' || q.lesson === lesson; }
function hashNum(s) { var h = 2166136261; for (var i = 0; i < s.length; i++) { h ^= s.charCodeAt(i); h = Math.imul(h, 16777619) >>> 0; } return h; }

// 某學生某課次的錯題資料：每題錯幾次、選過哪些答案、盲點；每個考點錯幾次
function wrongProfile(sid, lesson, Q) {
  var items = {}, kpw = {};
  wrongItems(sid).forEach(function (w) {
    var q = Q[w.qid];
    if (!q || !inLesson(q, lesson)) return;
    items[w.qid] = { qid: w.qid, count: w.count, status: w.status, hist: w.hist, attempts: w.attempts, added: w.addDay, choices: {}, traps: {} };
  });
  rows(SH.LOG).forEach(function (r) {
    if (str(r[1]) !== sid || str(r[7]) !== '錯') return;
    var qid = str(r[5]).split(/[#→]/)[0], q = Q[qid];
    if (!q || !inLesson(q, lesson)) return;
    if (q.kp) kpw[q.kp] = (kpw[q.kp] || 0) + 1;
    var it = items[qid];
    if (!it) return;
    var c = str(r[9]) || str(r[6]);
    if (c) it.choices[c] = (it.choices[c] || 0) + 1;
    if (str(r[12])) it.traps[str(r[12])] = (it.traps[str(r[12])] || 0) + 1;
  });
  var list = Object.keys(items).map(function (k) { return items[k]; })
    .sort(function (a, b) { return b.count - a.count || (a.qid < b.qid ? -1 : 1); });
  return { items: list, kpw: kpw };
}

// 出題：錯過的題目（錯越多次越優先）＋同考點的其他題（錯越多的考點越多題，沒做過的優先）
function buildDrill(sid, lesson, size, seed, Q) {
  var prof = wrongProfile(sid, lesson, Q), picked = [], used = {}, seen = {};
  rows(SH.LOG).forEach(function (r) { if (str(r[1]) === sid) seen[str(r[5]).split(/[#→]/)[0]] = 1; });
  var nOrig = Math.min(prof.items.length, Math.max(Math.ceil(size * 0.6), size - 10));
  prof.items.slice(0, nOrig).forEach(function (it) {
    picked.push({ q: Q[it.qid], type: 'wrong', count: it.count }); used[it.qid] = 1;
  });
  var kps = Object.keys(prof.kpw).sort(function (a, b) { return prof.kpw[b] - prof.kpw[a]; });
  var pools = {};
  kps.forEach(function (kp) {
    pools[kp] = Object.keys(Q).map(function (id) { return Q[id]; })
      .filter(function (x) { return x.kp === kp && inLesson(x, lesson) && !x.passage && !used[x.id]; })
      .sort(function (a, b) { return (seen[a.id] ? 1 : 0) - (seen[b.id] ? 1 : 0) || hashNum(seed + a.id) - hashNum(seed + b.id); });
  });
  var tickets = [];
  kps.forEach(function (kp) { for (var i = 0; i < Math.min(prof.kpw[kp], 4); i++) tickets.push(kp); });
  var guard = 0;
  while (picked.length < size && tickets.length && guard++ < 500) {
    var kp = tickets[guard % tickets.length], pool = pools[kp];
    while (pool.length && used[pool[0].id]) pool.shift();
    if (!pool.length) { tickets = tickets.filter(function (t) { return t !== kp; }); continue; }
    var x = pool.shift(); used[x.id] = 1;
    picked.push({ q: x, type: 'sibling', count: 0 });
  }
  picked.sort(function (a, b) { return hashNum(seed + a.q.id) - hashNum(seed + b.q.id); });
  return { list: picked, profile: prof };
}
// 考前特訓解鎖：某一課「只考這一課」的每日練習全部完成才開放；兩課一起要兩課都完成
function drillCounts(sid, days) {
  var Q = questionMap(), c = { L3: 0, L4: 0 };
  wrongItems(sid).forEach(function (w) { var q = Q[w.qid]; if (q && c[q.lesson] != null) c[q.lesson]++; });
  c.ALL = c.L3 + c.L4;
  if (!days) days = summaryDays(sid);
  var st = {}; days.forEach(function (d) { st[d.date] = d.status; });
  var need = { L3: [0, 0], L4: [0, 0] };
  schedule().forEach(function (d) {
    var ls = {};
    d.qids.forEach(function (id) { if (Q[id]) ls[Q[id].lesson] = 1; });
    var keys = Object.keys(ls);
    if (keys.length === 1 && need[keys[0]]) { need[keys[0]][1]++; if (st[d.date] === 'done') need[keys[0]][0]++; }
  });
  c.need = need;
  c.unlock = { L3: need.L3[1] > 0 && need.L3[0] >= need.L3[1], L4: need.L4[1] > 0 && need.L4[0] >= need.L4[1] };
  c.unlock.ALL = c.unlock.L3 && c.unlock.L4;
  return c;
}
function summaryDays(sid) {
  var log = studentLog(sid);
  return schedule().map(function (d) {
    var a = log.daily[d.date] || {}, n = 0;
    d.qids.forEach(function (id) { if (a[id]) n++; });
    return { date: d.date, status: n >= d.qids.length ? 'done' : 'todo' };
  });
}

API.getDrill = function (p) {
  var s = auth(p), Q = questionMap(), K = kpMap(), lesson = lessonKey(p.lesson);
  var dc = drillCounts(s.sid);
  if (!dc.unlock[lesson]) throw new Error(lesson === 'ALL' ? '第三課和第四課的每日練習都完成後，才能解鎖兩課一起的考前特訓' : LESSON_NAME[lesson] + '的每日練習全部完成後，才能解鎖考前特訓（目前 ' + dc.need[lesson][0] + ' / ' + dc.need[lesson][1] + ' 天）');
  var d = buildDrill(s.sid, lesson, 20, s.sid + '|' + today() + '|' + lesson, Q);
  if (!d.list.length) throw new Error(LESSON_NAME[lesson] + '還沒有錯過的題目，不用特訓喔！');
  var units = [];
  d.list.forEach(function (it) { if (units.indexOf(it.q.unit) < 0) units.push(it.q.unit); });
  return {
    ok: true, title: '考前特訓｜' + LESSON_NAME[lesson], lesson: lesson,
    units: unitsFor({ units: units, qids: [] }, Q), answered: {},
    questions: d.list.map(function (it) {
      var order = shuffledIdx(s.sid + '|drill|' + it.q.id + '|' + today());
      return { id: it.q.id, q: it.q.q, passage: it.q.passage || '', opts: order.map(function (i) { return it.q.opts[i]; }),
        unit: it.q.unit, kpName: (K[it.q.kp] || {}).name || '', drillType: it.type, wrongCount: it.count };
    })
  };
};

// 考前特訓作答：不計分；答錯一樣收進錯題本
function answerDrill(s, q, K, p) {
  var ss = SpreadsheetApp.getActive(), now = new Date();
  var right = q.opts['ABCD'.indexOf(q.ans)], choice = String(p.choiceText || '');
  var ci = q.opts.indexOf(choice), letter = ci >= 0 ? 'ABCD'.charAt(ci) : '';
  var correct = choice === right, trap = correct ? '' : trapOf(q.traps, letter);
  if (!correct) {
    var wsh = ss.getSheetByName(SH.WRONG), wv = wsh.getDataRange().getValues(), wrow = -1;
    for (var i = 1; i < wv.length; i++) if (str(wv[i][0]) === s.sid && str(wv[i][1]) === q.id && str(wv[i][3]) === '待訂正') { wrow = i + 1; break; }
    if (wrow > 0) { wsh.getRange(wrow, 3).setValue((Number(wv[wrow - 1][2]) || 1) + 1); wsh.getRange(wrow, 6).setValue(now); }
    else wsh.appendRow([s.sid, q.id, 1, '待訂正', now, now, 0, '']);
  }
  ss.getSheetByName(SH.LOG).appendRow([now, s.sid, s.name, 'drill', '', q.id, letter, correct ? '對' : '錯', 0, choice, right, q.kp, trap]);
  recordAnalysis(s, q, letter, correct, trap, q.id, choice);
  SpreadsheetApp.flush();
  var sum = refreshStudent(s), k = K[q.kp] || {};
  return { ok: true, correct: correct, ansText: right, exp: q.exp, gained: 0, points: sum.points, kpName: k.name || '', kpTip: k.tip || '', trap: trap };
}

// ---------- 老師工具：在「老師工具」工作表選學生、課次，勾選「產生」 ----------
var TOOL = '老師工具', REPORT = '錯題報告', PAPER = '考前練習卷', KEYS = '練習卷解答';
function setupTeacherTool() {
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(TOOL) || ss.insertSheet(TOOL);
  sh.getRange(1, 1, 8, 2).setValues([
    ['📊 老師工具：錯題報告＋考前練習卷', ''],
    ['', ''],
    ['學生（登入名稱）', ''],
    ['課次', '第三課'],
    ['練習卷題數', 30],
    ['✅ 勾選這格就會產生 →', false],
    ['狀態', '選好學生和課次後，勾選上面那格'],
    ['說明', '產生後看「錯題報告」「考前練習卷」「練習卷解答」三個工作表；要列印就用「檔案 → 列印」選目前工作表。每次產生會覆蓋上一次的內容。']
  ]);
  sh.getRange(1, 1).setFontWeight('bold').setFontSize(14);
  sh.getRange(3, 1, 6, 1).setFontWeight('bold');
  sh.getRange(3, 2, 4, 1).setBackground('#fff8e1');
  sh.setColumnWidth(1, 190); sh.setColumnWidth(2, 420);
  sh.getRange(8, 2).setWrap(true);
  var stu = ss.getSheetByName(SH.STU);
  sh.getRange(3, 2).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(stu.getRange('A2:A'), true).build());
  sh.getRange(4, 2).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['第三課', '第四課', '兩課一起'], true).build());
  sh.getRange(6, 2).insertCheckboxes();
  var first = rows(SH.STU).filter(function (r) { return str(r[0]); })[0];
  if (first) sh.getRange(3, 2).setValue(str(first[0]));
  return true;
}
function onEdit(e) {
  try {
    var r = e && e.range;
    if (!r || r.getSheet().getName() !== TOOL || r.getRow() !== 6 || r.getColumn() !== 2 || r.getValue() !== true) return;
    var sh = r.getSheet();
    sh.getRange(7, 2).setValue('產生中…');
    var msg = teacherReportFromTool();
    sh.getRange(7, 2).setValue(msg);
  } catch (err) {
    if (e && e.range) e.range.getSheet().getRange(7, 2).setValue('❌ ' + (err.message || err));
  } finally {
    if (e && e.range && e.range.getSheet().getName() === TOOL && e.range.getRow() === 6 && e.range.getColumn() === 2) e.range.setValue(false);
  }
}
function teacherReportFromTool() {
  var sh = SpreadsheetApp.getActive().getSheetByName(TOOL), v = sh.getRange(3, 2, 3, 1).getValues();
  return teacherReport(str(v[0][0]), str(v[1][0]), Number(v[2][0]) || 30);
}
function plain(html) {
  return str(html).replace(/<br\s*\/?>/gi, '\n').replace(/<[^>]+>/g, '').replace(/&nbsp;/g, ' ').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
}
function passagePlain(p) {
  if (!p) return '';
  var m = p.match(/^\[\[(\d+)\]\]/), cur = m ? m[1] : '';
  return p.replace(/^\[\[\d+\]\]/, '').replace(/\{(\d+)\}/g, function (_, k) { return k === cur ? '【 ' + k + ' 】' : '( ' + k + ' )'; });
}
function teacherReport(who, lessonLabel, size) {
  var ss = SpreadsheetApp.getActive(), Q = questionMap(), K = kpMap(), lesson = lessonKey(lessonLabel);
  var s = auth({ sid: who });
  var stamp = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm');
  // ---- 考點總覽（這一課所有作答） ----
  var B = {}, order = [];
  rows(SH.LOG).forEach(function (r) {
    if (str(r[1]) !== s.sid) return;
    var qlabel = str(r[5]), q = Q[qlabel.split(/[#→]/)[0]];
    if (!q || !q.kp || !inLesson(q, lesson)) return;
    if (!B[q.kp]) { B[q.kp] = [s.sid, s.name, q.kp, (K[q.kp] || {}).name || q.kp, 0, 0, 0, '', '']; order.push(q.kp); }
    B[q.kp] = blindRow(B[q.kp], str(r[7]) === '對', str(r[12]), qlabel, str(r[9]) || str(r[6]));
  });
  order.sort(function (a, b) { return B[b][6] - B[a][6] || B[b][5] - B[a][5]; });
  var prof = wrongProfile(s.sid, lesson, Q);
  var out = [];
  out.push(['錯題分析報告｜' + s.name + '｜' + LESSON_NAME[lesson] + '｜產生時間 ' + stamp, '', '', '', '', '', '', '', '', '']);
  out.push(['', '', '', '', '', '', '', '', '', '']);
  out.push(['一、考點總覽（錯誤率高的在上面）', '', '', '', '', '', '', '', '', '']);
  out.push(['考點', '作答數', '答錯數', '錯誤率', '盲點統計', '', '', '', '', '']);
  if (!order.length) out.push(['（這一課還沒有作答紀錄）', '', '', '', '', '', '', '', '', '']);
  order.forEach(function (kp) { var b = B[kp]; out.push([b[3], b[4], b[5], b[6], b[7], '', '', '', '', '']); });
  out.push(['', '', '', '', '', '', '', '', '', '']);
  out.push(['二、錯題清單（共 ' + prof.items.length + ' 題，錯越多次的在上面）', '', '', '', '', '', '', '', '', '']);
  var h2 = out.length + 1;
  out.push(['題目ID', '考點', '題目', '她選過的答案（次數）', '正確答案', '錯誤次數', '狀態', '回收紀錄', '盲點', '詳解']);
  if (!prof.items.length) out.push(['（這一課沒有錯題）', '', '', '', '', '', '', '', '', '']);
  prof.items.forEach(function (it) {
    var q = Q[it.qid];
    out.push([it.qid, (K[q.kp] || {}).name || q.kp, (q.passage ? '［克漏字］' + passagePlain(q.passage) + '\n' : '') + q.q,
      Object.keys(it.choices).map(function (c) { return c + '×' + it.choices[c]; }).join('\n'),
      '(' + q.ans + ') ' + q.opts['ABCD'.indexOf(q.ans)], it.count, it.status, it.hist,
      Object.keys(it.traps).join('\n'), plain(q.exp)]);
  });
  var rep = ss.getSheetByName(REPORT) || ss.insertSheet(REPORT);
  rep.clear();
  rep.getRange(1, 1, out.length, 10).setValues(out).setWrap(true).setVerticalAlignment('top');
  rep.getRange(1, 1).setFontWeight('bold').setFontSize(13);
  var h1 = 4;
  rep.getRange(3, 1).setFontWeight('bold'); rep.getRange(h2 - 1, 1).setFontWeight('bold');
  rep.getRange(h1, 1, 1, 5).setFontWeight('bold').setBackground('#e3f1ef');
  rep.getRange(h2, 1, 1, 10).setFontWeight('bold').setBackground('#e3f1ef');
  if (order.length) rep.getRange(h1 + 1, 4, order.length, 1).setNumberFormat('0%');
  [80, 160, 360, 160, 140, 60, 60, 70, 200, 320].forEach(function (w, i) { rep.setColumnWidth(i + 1, w); });

  // ---- 考前練習卷＋解答 ----
  var d = buildDrill(s.sid, lesson, Math.max(5, Math.min(size, 60)), s.sid + '|paper|' + stamp, Q);
  var paper = [['考前練習卷｜' + LESSON_NAME[lesson] + '　　姓名：' + s.name + '　　日期：＿＿＿＿　　得分：＿＿＿＿', '']];
  var keys = [['考前練習卷解答｜' + s.name + '｜' + LESSON_NAME[lesson] + '｜' + stamp, '', '', ''], ['題號', '答案', '考點', '詳解']];
  d.list.forEach(function (it, i) {
    var q = it.q, ord = shuffledIdx(s.sid + '|paper|' + q.id + '|' + stamp), opts = ord.map(function (j) { return q.opts[j]; });
    var ansL = 'ABCD'.charAt(opts.indexOf(q.opts['ABCD'.indexOf(q.ans)]));
    var body = (q.passage ? passagePlain(q.passage) + '\n\n' : '') + q.q.replace(/_{3,}/g, '＿＿＿＿') + '\n' +
      opts.map(function (o, j) { return '(' + 'ABCD'.charAt(j) + ') ' + o; }).join('　　');
    paper.push(['(　　) ' + (i + 1) + '.', body]);
    keys.push([i + 1, ansL, (K[q.kp] || {}).name || q.kp, plain(q.exp) + (it.type === 'wrong' ? '　［她錯過這題 ' + it.count + ' 次］' : '')]);
  });
  var ps = ss.getSheetByName(PAPER) || ss.insertSheet(PAPER);
  ps.clear();
  ps.getRange(1, 1, paper.length, 2).setValues(paper).setWrap(true).setVerticalAlignment('top');
  ps.getRange(1, 1).setFontWeight('bold').setFontSize(13);
  ps.setColumnWidth(1, 90); ps.setColumnWidth(2, 720);
  var ks = ss.getSheetByName(KEYS) || ss.insertSheet(KEYS);
  ks.clear();
  ks.getRange(1, 1, keys.length, 4).setValues(keys).setWrap(true).setVerticalAlignment('top');
  ks.getRange(1, 1).setFontWeight('bold');
  ks.getRange(2, 1, 1, 4).setFontWeight('bold').setBackground('#e3f1ef');
  [50, 50, 180, 560].forEach(function (w, i) { ks.setColumnWidth(i + 1, w); });
  return '✅ ' + stamp + ' 已產生：' + s.name + '／' + LESSON_NAME[lesson] + '，錯題 ' + prof.items.length + ' 題、練習卷 ' + d.list.length + ' 題';
}
