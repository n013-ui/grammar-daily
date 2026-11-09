(function () {
  "use strict";

  var CFG = window.APP_CONFIG || {};
  var DEMO = /[?&]demo=1/.test(location.search) || !CFG.API_URL;   // 網址加 ?demo=1 可強制試看模式
  var $ = function (id) { return document.getElementById(id); };
  var WEEK = ["日", "一", "二", "三", "四", "五", "六"];

  // ---------- 小工具 ----------
  function esc(s) {
    return String(s).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; });
  }
  function qHtml(s) { return esc(s).replace(/_{3,}/g, '<span class="blank"></span>').replace(/\n/g, "<br>"); }
  function parseDate(s) { var p = s.split("-"); return new Date(+p[0], +p[1] - 1, +p[2]); }
  function fmtDate(d) { return d.getFullYear() + "-" + ("0" + (d.getMonth() + 1)).slice(-2) + "-" + ("0" + d.getDate()).slice(-2); }
  function niceDate(s) { var d = parseDate(s); return (d.getMonth() + 1) + "/" + d.getDate() + "（" + WEEK[d.getDay()] + "）"; }
  function store(k, v) { try { if (v === undefined) return JSON.parse(localStorage.getItem(k)); localStorage.setItem(k, JSON.stringify(v)); } catch (e) { return null; } }
  function toast(t) { var el = $("toast"); el.hidden = true; void el.offsetWidth; el.textContent = t; el.hidden = false; clearTimeout(toast._t); toast._t = setTimeout(function () { el.hidden = true; }, 950); }
  function loading(on) { $("loading").hidden = !on; }
  function histHtml(h) { return (h || "").split("").map(function (c) { return c === "✓" ? '<span class="h ok">✓</span>' : '<span class="h bad">✗</span>'; }).join(""); }
  function shuffledIdx(seedStr) {   // 與 GAS 相同的固定亂數
    var h = 2166136261;
    for (var i = 0; i < seedStr.length; i++) { h ^= seedStr.charCodeAt(i); h = Math.imul(h, 16777619) >>> 0; }
    var a = [0, 1, 2, 3];
    for (var j = 3; j > 0; j--) { h = (Math.imul(h, 1103515245) + 12345) >>> 0; var r = h % (j + 1); var t = a[j]; a[j] = a[r]; a[r] = t; }
    return a;
  }

  // ---------- 後端：正式（GAS） ----------
  var LiveApi = {
    call: function (action, data) {
      data = data || {};
      data.action = action;
      if (session) { data.sid = session.sid; data.name = session.name; }
      return fetch(CFG.API_URL, { method: "POST", body: JSON.stringify(data) })
        .then(function (r) { return r.json(); })
        .then(function (res) { if (!res.ok) throw new Error(res.error || "發生錯誤"); return res; });
    }
  };

  // ---------- 後端：試看模式（瀏覽器內模擬，規則與 GAS 相同） ----------
  var DemoApi = (function () {
    var D, KEY = "grammarDemo_v2";
    function db() { return store(KEY) || { students: {} }; }
    function save(x) { store(KEY, x); }
    function today() { return $("demoDate").value; }
    function qmap() { var m = {}; D.questions.forEach(function (q) { m[q.id] = q; }); return m; }
    function umap() { var m = {}; D.units.forEach(function (u) { m[u.id] = u; }); return m; }
    function vars(id) { return ((D.variants || {})[id] || []).map(function (v) { return { q: v[0], opts: v[1], ans: v[2], exp: v[3] }; }); }
    function pick(q, k, exact) {
      var V = vars(q.id);
      if (!V.length || k < 0) return { vk: -1, q: q.q, opts: q.opts, ans: q.ans, exp: q.exp };
      var i = exact ? k : k % (V.length + 1), v = V[i];
      return v ? { vk: i, q: v.q, opts: v.opts, ans: v.ans, exp: v.exp } : { vk: -1, q: q.q, opts: q.opts, ans: q.ans, exp: q.exp };
    }
    function stu(x) {
      if (!x.students[session.sid]) x.students[session.sid] = { name: session.name, points: 0, daily: {}, wrong: [] };
      return x.students[session.sid];
    }
    function pub(q) { return { id: q.id, q: q.q, opts: q.opts, src: q.src, unit: q.unit }; }
    function unitsFor(day, Q) {
      var ids = day.units && day.units.length ? day.units.slice() : [];
      if (!ids.length) day.qids.forEach(function (id) { var u = Q[id] && Q[id].unit; if (u && ids.indexOf(u) < 0) ids.push(u); });
      var U = umap();
      return ids.sort().map(function (id) { return U[id]; }).filter(Boolean);
    }
    function pendingRow(s, qid) { return s.wrong.filter(function (w) { return w.qid === qid && w.status === "待訂正"; })[0]; }
    function summary(s) {
      var t = today();
      var days = D.schedule.map(function (d) {
        var a = s.daily[d.date] || {}, n = 0, sc = 0;
        d.qids.forEach(function (id) { if (a[id]) { n++; if (a[id].correct) sc++; } });
        var st = d.date > t ? "locked" : n >= d.qids.length ? "done" : n > 0 ? "partial" : "todo";
        return { date: d.date, title: d.title, total: d.qids.length, answered: n, score: sc, status: st };
      });
      var pend = s.wrong.filter(function (w) { return w.status === "待訂正"; });
      return {
        ok: true, today: t, points: s.points, days: days, wrongCount: pend.length,
        reviewReady: pend.filter(function (w) { return w.lastDay < t; }).length,
        reviewWaiting: pend.filter(function (w) { return w.lastDay >= t; }).length,
        fixedCount: s.wrong.filter(function (w) { return w.status === "已訂正"; }).length,
        reviewCount: D.settings.reviewCount,
        doneDays: days.filter(function (d) { return d.status === "done"; }).length,
        student: { sid: session.sid, name: s.name }
      };
    }
    var api = {
      call: function (action, data) {
        return new Promise(function (resolve, reject) {
          setTimeout(function () { try { resolve(api[action](data || {})); } catch (e) { reject(e); } }, 120);
        });
      },
      login: function () { var x = db(); var s = stu(x); save(x); return summary(s); },
      summary: function () { return summary(stu(db())); },
      getDay: function (p) {
        var day = D.schedule.filter(function (d) { return d.date === p.date; })[0];
        if (!day) throw new Error("這一天沒有排練習");
        if (day.date > today()) throw new Error("這一天還沒開放");
        var Q = qmap(), s = stu(db()), a = s.daily[day.date] || {}, answered = {};
        day.qids.forEach(function (id) { if (a[id]) answered[id] = { choice: a[id].choice, correct: a[id].correct, ans: Q[id].ans, exp: Q[id].exp }; });
        return { ok: true, date: day.date, title: day.title, units: unitsFor(day, Q),
          questions: day.qids.map(function (id) { return pub(Q[id]); }), answered: answered };
      },
      getReview: function () {
        var Q = qmap(), s = stu(db()), t = today();
        var items = s.wrong.filter(function (w) { return w.status === "待訂正" && w.lastDay < t; }).slice(0, D.settings.reviewCount);
        return { ok: true, title: "錯題回收區", units: unitsFor({ units: [], qids: items.map(function (w) { return w.qid; }) }, Q),
          answered: {}, questions: items.map(function (w) {
            var src = pick(Q[w.qid], w.attempts), order = shuffledIdx(session.sid + "|" + w.qid + "|" + w.attempts);
            return { id: w.qid, vk: src.vk, q: src.q, opts: order.map(function (i) { return src.opts[i]; }), unit: Q[w.qid].unit,
              src: src.vk >= 0 ? "變化題" : "原題", wrongCount: w.count, attempts: w.attempts, hist: w.hist };
          }) };
      },
      wrongBook: function () {
        var Q = qmap(), s = stu(db()), t = today();
        var items = s.wrong.map(function (w) {
          return { qid: w.qid, q: Q[w.qid].q, count: w.count, status: w.status, added: w.addDay, attempts: w.attempts, hist: w.hist,
            ready: w.status === "待訂正" && w.lastDay < t };
        });
        items.sort(function (a, b) { return (a.status === "待訂正" ? 0 : 1) - (b.status === "待訂正" ? 0 : 1); });
        return { ok: true, items: items };
      },
      answer: function (p) {
        var Q = qmap(), q = Q[p.qid]; if (!q) throw new Error("找不到題目");
        var x = db(), s = stu(x), t = today(), w = pendingRow(s, p.qid);
        if (p.mode === "review") {
          if (!w) throw new Error("這題已經不在錯題本裡了");
          if (w.lastDay >= t) throw new Error("這題今天剛錯過，明天才能練");
          var src = pick(q, Number(p.vk), true), right = src.opts["ABCD".indexOf(src.ans)], ok = p.choiceText === right, g = 0;
          w.attempts++; w.hist += ok ? "✓" : "✗"; w.lastDay = t;
          if (ok) { w.status = "已訂正"; g = 1; } else w.count++;
          s.points += g; save(x);
          return { ok: true, correct: ok, ansText: right, exp: src.exp, gained: g, points: s.points, wrongCount: w.count, hist: w.hist };
        }
        var correct = p.choice === q.ans, gained = 0;
        if (p.mode === "daily") {
          var a = s.daily[p.date] = s.daily[p.date] || {};
          if (a[p.qid]) return { ok: true, correct: a[p.qid].correct, ans: q.ans, exp: q.exp, gained: 0, points: s.points, repeat: true };
          a[p.qid] = { choice: p.choice, correct: correct };
          if (correct) gained = 1;
        }
        if (!correct) {
          if (w) { w.count++; w.lastDay = t; }
          else s.wrong.push({ qid: p.qid, count: 1, status: "待訂正", addDay: t, lastDay: t, attempts: 0, hist: "" });
        }
        s.points += gained; save(x);
        return { ok: true, correct: correct, ans: q.ans, exp: q.exp, gained: gained, points: s.points };
      },
      init: function () {
        D = window.GRAMMAR_DATA;
        var d = $("demoDate"), first = D.schedule[0].date, last = D.schedule[D.schedule.length - 1].date;
        var t = fmtDate(new Date());
        d.min = first.slice(0, 8) + "01"; d.max = last.slice(0, 8) + "30";
        d.value = store("grammarDemoDate") || (t < first ? first : t > last ? last : t);
        d.addEventListener("change", function () { store("grammarDemoDate", d.value); if (session) goHome(); });
        $("demoReset").addEventListener("click", function () {
          if (!confirm("清除這台電腦上所有試看模式的作答紀錄？")) return;
          store(KEY, { students: {} }); if (session) goHome();
        });
        $("demoBar").hidden = false;
      }
    };
    return api;
  })();

  var Api = DEMO ? DemoApi : LiveApi;

  // ---------- 狀態 ----------
  var SKEY = DEMO ? "grammarSessionDemo" : "grammarSession";
  var session = store(SKEY);
  var home = null;   // 最近一次 summary
  var quiz = null;   // { mode, date, title, units, questions, idx, results }

  function show(v) {
    ["vLogin", "vHome", "vNotes", "vQuiz", "vResult", "vBook"].forEach(function (id) { $(id).hidden = id !== v; });
    $("who").hidden = !session || v === "vLogin";
    window.scrollTo(0, 0);
  }
  function setPoints(n) { if (typeof n === "number") $("pointsNum").textContent = n; }
  function fail(e) { loading(false); alert(e.message || e); }

  // ---------- 登入 ----------
  $("loginForm").addEventListener("submit", function (e) {
    e.preventDefault();
    var sid = $("inSid").value.trim(), name = $("inName").value.trim();
    if (!sid || !name) return;
    session = { sid: sid, name: name };
    $("loginErr").textContent = "";
    loading(true);
    Api.call("login").then(function (res) {
      loading(false);
      session.name = res.student.name;
      store(SKEY, session);
      renderHome(res);
    }).catch(function (err) { loading(false); session = null; $("loginErr").textContent = err.message; });
  });
  $("logoutBtn").addEventListener("click", function () { session = null; store(SKEY, null); $("inSid").value = ""; $("inName").value = ""; show("vLogin"); });

  // ---------- 首頁 ----------
  function goHome() {
    loading(true);
    Api.call("summary").then(function (res) { loading(false); renderHome(res); }).catch(fail);
  }
  function renderHome(res) {
    home = res;
    if (res.title) { $("appTitle").textContent = res.title; document.title = res.title; }
    $("whoName").textContent = res.student.name;
    setPoints(res.points);
    $("statDone").textContent = res.doneDays + " / " + res.days.length;
    $("statFixed").textContent = res.fixedCount;

    var t = res.today, todayDay = res.days.filter(function (d) { return d.date === t; })[0];
    var h = "", wd = parseDate(t).getDay();
    if (todayDay) {
      h += '<div class="eyebrow">今天 ' + niceDate(t) + " · 第 " + (res.days.indexOf(todayDay) + 1) + " 天</div>";
      h += "<h2>" + esc(todayDay.title) + "</h2>";
      h += statusBadge(todayDay);
      h += '<div class="btn-row">' + dayButtons(todayDay) + "</div>";
    } else {
      var first = res.days[0], last = res.days[res.days.length - 1];
      h += '<div class="eyebrow">今天 ' + niceDate(t) + "</div>";
      if (first && t < first.date) h += "<h2>練習從 " + niceDate(first.date) + " 開始</h2><p class=\"muted\">到時候每天 10 題，答對一題得 1 點！</p>";
      else if (last && t > last.date) h += "<h2>本月練習已全部結束 🎉</h2><p class=\"muted\">可以到錯題回收區把錯題清光，或從月曆回顧題目。</p>";
      else if (wd === 0 || wd === 6) h += "<h2>週末沒有每日十題 ☕</h2><p class=\"muted\">可以到下面的錯題回收區練習，或補做還沒完成的天數。</p>";
      else h += "<h2>今天沒有排練習</h2><p class=\"muted\">可以到錯題回收區練習，或補做還沒完成的天數。</p>";
    }
    $("todayCard").innerHTML = h;

    // 錯題回收區
    var r = "<h3>♻️ 錯題回收區</h3>";
    r += '<div class="review-nums"><span>現在可練 <b>' + res.reviewReady + "</b> 題</span><span>明天開放 <b>" + res.reviewWaiting +
      "</b> 題</span><span>已訂正 <b>" + res.fixedCount + "</b> 題</span></div>";
    if (res.reviewReady > 0) {
      r += '<p class="muted" style="margin:6px 0 0">以前答錯的題目會換個句子考同樣的觀念，答對就訂正成功，+1 點。</p>';
      r += '<div class="btn-row"><button class="btn review big" data-review="1">開始錯題回收（' + Math.min(res.reviewCount || 10, res.reviewReady) +
        ' 題）</button><button class="btn big" data-book="1">看錯題本</button></div>';
    } else {
      r += '<p class="muted" style="margin:6px 0 0">' + (res.reviewWaiting > 0 ? "今天答錯的題目，明天就能來這裡練習。" : res.fixedCount > 0 ? "錯題都訂正完了，太厲害了！" : "答錯的題目會收進這裡，隔天就能練習。") + "</p>";
      if (res.wrongCount + res.fixedCount > 0) r += '<button class="btn big" data-book="1">看錯題本</button>';
    }
    $("reviewCard").innerHTML = r;

    // 補做
    var miss = res.days.filter(function (d) { return d.date < t && (d.status === "todo" || d.status === "partial"); });
    var mk = $("makeupCard");
    mk.hidden = !miss.length;
    if (miss.length) {
      mk.innerHTML = '<h3 class="card-title">還沒完成的練習（可以補做）</h3>' + miss.map(function (d) {
        return '<div class="makeup-item"><div><b>' + niceDate(d.date) + "</b> " + esc(d.title) +
          (d.status === "partial" ? '<br><small class="muted">已做 ' + d.answered + " / " + d.total + " 題</small>" : "") +
          '</div><button class="btn" data-day="' + d.date + '">' + (d.status === "partial" ? "繼續" : "補做") + "</button></div>";
      }).join("");
    }
    renderCal(res);
    show("vHome");
  }
  function statusBadge(d) {
    if (d.status === "done") return '<span class="status done">已完成 · 答對 ' + d.score + " / " + d.total + "</span>";
    if (d.status === "partial") return '<span class="status partial">做到第 ' + d.answered + " 題</span>";
    return '<span class="status todo">尚未開始 · ' + d.total + " 題</span>";
  }
  function dayButtons(d) {
    if (d.status === "done") return '<button class="btn big" data-result="' + d.date + '">看今天的解析</button><button class="btn primary big" data-redo="' + d.date + '">再練一次（不計分）</button>';
    return '<button class="btn primary big" data-day="' + d.date + '">' + (d.status === "partial" ? "繼續練習 →" : "看文法說明，開始練習 →") + "</button>";
  }
  function renderCal(res) {
    var byDate = {}; res.days.forEach(function (d) { byDate[d.date] = d; });
    var base = parseDate((res.days[0] || { date: res.today }).date);
    var y = base.getFullYear(), m = base.getMonth();
    var first = new Date(y, m, 1), n = new Date(y, m + 1, 0).getDate(), h = "";
    for (var i = 0; i < first.getDay(); i++) h += '<div class="day blank"></div>';
    for (var dd = 1; dd <= n; dd++) {
      var ds = fmtDate(new Date(y, m, dd)), d = byDate[ds];
      var cls = "day", sub = "", attr = "";
      if (d) {
        cls += " " + d.status;
        sub = d.status === "done" ? d.score + "/" + d.total : d.status === "partial" ? d.answered + "/" + d.total : d.status === "locked" ? "" : "待做";
        if (d.status !== "locked") attr = d.status === "done" ? ' data-result="' + ds + '"' : ' data-day="' + ds + '"';
      } else cls += " locked";
      if (ds === res.today) cls += " today";
      h += attr ? '<button class="' + cls + '"' + attr + ">" + dd + "<small>" + sub + "</small></button>"
                : '<div class="' + cls + '">' + dd + "<small>" + sub + "</small></div>";
    }
    $("cal").innerHTML = h;
  }

  // 按鈕（事件委派）
  document.addEventListener("click", function (e) {
    var b = e.target.closest("[data-day],[data-redo],[data-result],[data-review],[data-book],[data-go]");
    if (!b) return;
    if (b.dataset.go === "home") return goHome();
    if (b.dataset.day) return openDay(b.dataset.day, "daily");
    if (b.dataset.redo) return openDay(b.dataset.redo, "practice");
    if (b.dataset.result) return openResult(b.dataset.result);
    if (b.dataset.review) return openReview();
    if (b.dataset.book) return openBook();
  });

  // ---------- 錯題本 ----------
  function openBook() {
    loading(true);
    Api.call("wrongBook").then(function (res) {
      loading(false);
      $("bookList").innerHTML = res.items.length ? res.items.map(function (w) {
        var st = w.status === "已訂正" ? '<span class="st fixed">已訂正</span>'
          : w.ready ? '<span class="st ready">可以練習</span>' : '<span class="st pending">明天開放</span>';
        var again = w.hist && w.hist.slice(-1) === "✗" && w.status !== "已訂正" ? '<span class="badge bad">上次又錯了</span>' : "";
        return '<div class="book-item' + (w.status === "已訂正" ? " fixed" : "") + '"><div class="bq">' + qHtml(w.q) + '</div><div class="bmeta">' + st + again +
          "<span>錯 " + w.count + " 次</span>" + (w.attempts ? "<span>回收紀錄 " + histHtml(w.hist) + "</span>" : "<span>還沒回收練習過</span>") +
          '<span class="muted">' + (w.added ? niceDate(w.added) + " 加入" : "") + "</span></div></div>";
      }).join("") : '<p class="muted">錯題本是空的！</p>';
      show("vBook");
    }).catch(fail);
  }

  // ---------- 開始練習 ----------
  function openDay(date, mode) {
    loading(true);
    Api.call("getDay", { date: date }).then(function (res) {
      loading(false);
      quiz = { mode: mode, date: res.date, title: res.title, units: res.units, questions: res.questions,
        results: mode === "daily" ? res.answered : {}, idx: 0 };
      renderNotes();
    }).catch(fail);
  }
  function openResult(date) {
    loading(true);
    Api.call("getDay", { date: date }).then(function (res) {
      loading(false);
      quiz = { mode: "daily", date: res.date, title: res.title, units: res.units, questions: res.questions, results: res.answered, idx: res.questions.length };
      renderResult(true);
    }).catch(fail);
  }
  function openReview() {
    loading(true);
    Api.call("getReview").then(function (res) {
      loading(false);
      if (!res.questions.length) { alert("現在沒有可以練習的錯題！"); return goHome(); }
      quiz = { mode: "review", date: home.today, title: res.title, units: res.units, questions: res.questions, results: {}, idx: 0 };
      startQuiz();
    }).catch(fail);
  }
  function unitsHtml(units) {
    return units.map(function (u) { return '<h3 class="unit-h">' + esc(u.lesson) + "｜" + esc(u.title) + "</h3>" + u.html; }).join("");
  }
  function renderNotes() {
    var mixed = /綜合/.test(quiz.title);
    $("notesEyebrow").textContent = niceDate(quiz.date) + (quiz.mode === "practice" ? " · 再練一次（不計分）" : "") + " · 先讀文法說明";
    $("notesTitle").textContent = quiz.title;
    $("notesBody").innerHTML = (mixed ? '<p class="tip">今天是綜合練習，題目來自下面這些文法重點，先快速複習一下。</p>' : "") + unitsHtml(quiz.units);
    var done = Object.keys(quiz.results).length;
    $("startQuizBtn").textContent = done ? "繼續練習（第 " + (done + 1) + " 題起）→" : "我讀完了，開始練習 " + quiz.questions.length + " 題 →";
    show("vNotes");
  }
  $("startQuizBtn").addEventListener("click", startQuiz);

  // ---------- 作答 ----------
  function startQuiz() {
    quiz.idx = 0;
    while (quiz.idx < quiz.questions.length && quiz.results[quiz.questions[quiz.idx].id]) quiz.idx++;
    if (quiz.idx >= quiz.questions.length) return renderResult();
    show("vQuiz");
    renderQ();
  }
  function renderQ() {
    var q = quiz.questions[quiz.idx], n = quiz.questions.length, rv = quiz.mode === "review";
    $("quizCount").textContent = (rv ? "錯題 " : "第 ") + (quiz.idx + 1) + " / " + n + (rv ? "" : " 題");
    $("progressBar").style.width = (quiz.idx / n * 100) + "%";
    $("progressBar").style.background = rv ? "var(--review)" : "";
    var meta = "";
    if (rv) {
      meta = '<div class="qbadges"><span class="badge">🔁 第 ' + (q.attempts + 1) + " 次回收練習</span>" +
        (q.src === "變化題" ? '<span class="badge">換句話考同觀念</span>' : '<span class="badge">原題再挑戰（選項已換位置）</span>') +
        '<span class="badge bad">已錯 ' + q.wrongCount + " 次</span>" +
        (q.hist && q.hist.slice(-1) === "✗" ? '<span class="badge bad">上次又錯了</span>' : "") + "</div>";
    }
    $("qSrc").textContent = rv ? "錯題回收 · " + q.id.slice(0, 2) : q.id.slice(0, 2) + " · " + q.src;
    $("qText").innerHTML = meta + qHtml(q.q);
    $("qOpts").innerHTML = q.opts.map(function (o, i) {
      var L = "ABCD"[i];
      return '<button class="opt" data-c="' + L + '"><span class="letter">' + L + "</span><span>" + esc(o) + "</span></button>";
    }).join("");
    $("qFeedback").hidden = true;
    $("nextBtn").hidden = true;
    $("peekNotes").hidden = !quiz.units.length;
  }
  $("qOpts").addEventListener("click", function (e) {
    var b = e.target.closest(".opt"); if (!b || b.disabled) return;
    var q = quiz.questions[quiz.idx], choice = b.dataset.c, rv = quiz.mode === "review";
    var all = $("qOpts").querySelectorAll(".opt");
    all.forEach(function (x) { x.disabled = true; });
    b.style.borderColor = "var(--primary)";
    var payload = { mode: quiz.mode, date: quiz.date, qid: q.id, choice: choice };
    if (rv) { payload.vk = q.vk; payload.choiceText = q.opts["ABCD".indexOf(choice)]; }
    Api.call("answer", payload).then(function (res) {
      var ansL = res.ansText != null ? "ABCD"[q.opts.indexOf(res.ansText)] : res.ans;
      quiz.results[q.id] = { choice: choice, correct: res.correct, ans: ansL, exp: res.exp };
      b.style.borderColor = "";
      all.forEach(function (x) {
        if (x.dataset.c === ansL) x.classList.add("correct");
        else if (x.dataset.c === choice) x.classList.add("wrong");
      });
      var fb = $("qFeedback");
      fb.className = "feedback " + (res.correct ? "ok" : "bad");
      var head = res.correct ? (rv ? "訂正成功！" : "答對了！") : (rv ? "又錯了！" : "答錯了，") + "正確答案是 (" + ansL + ")";
      if (res.gained) head += "　+" + res.gained + " 點";
      var sub = "";
      if (!res.correct && !rv) sub = "已收進錯題本，明天起可以到錯題回收區練習";
      if (!res.correct && rv) sub = "這題已錯 " + res.wrongCount + " 次，會留在錯題本，明天再換題練一次";
      if (res.correct && rv) sub = "已從錯題本移除。回收紀錄：" + res.hist;
      fb.innerHTML = '<div class="fb-head">' + head + (sub ? '<div style="font-size:13px;font-weight:500;color:var(--muted)">' + sub + "</div>" : "") +
        '</div><div class="exp">💬 ' + res.exp + "</div>";
      fb.hidden = false;
      if (res.gained) toast("+" + res.gained + " 點 ⭐");
      setPoints(res.points);
      var last = quiz.idx >= quiz.questions.length - 1;
      $("nextBtn").textContent = last ? "看結果 →" : "下一題 →";
      $("nextBtn").hidden = false;
      $("progressBar").style.width = ((quiz.idx + 1) / quiz.questions.length * 100) + "%";
      fb.scrollIntoView({ behavior: "smooth", block: "nearest" });
    }).catch(function (err) {
      all.forEach(function (x) { x.disabled = false; });
      b.style.borderColor = "";
      alert("送出失敗：" + err.message);
    });
  });
  $("nextBtn").addEventListener("click", function () {
    quiz.idx++;
    if (quiz.idx >= quiz.questions.length) return renderResult();
    renderQ();
    window.scrollTo(0, 0);
  });
  $("quizBack").addEventListener("click", goHome);
  $("peekNotes").addEventListener("click", function () { $("modalBody").innerHTML = unitsHtml(quiz.units); $("notesModal").hidden = false; });
  $("closeNotes").addEventListener("click", function () { $("notesModal").hidden = true; });
  $("notesModal").addEventListener("click", function (e) { if (e.target === this) this.hidden = true; });

  // ---------- 結果 ----------
  function renderResult(viewOnly) {
    var total = quiz.questions.length, score = 0, rv = quiz.mode === "review";
    quiz.questions.forEach(function (q) { var r = quiz.results[q.id]; if (r && r.correct) score++; });
    $("resEyebrow").textContent = rv ? "錯題回收結果" : niceDate(quiz.date) + " · " + quiz.title;
    $("resScore").textContent = score;
    $("resTotal").textContent = total;
    document.querySelector(".score-ring").style.setProperty("--pct", (score / total * 100) + "%");
    var r = score / total;
    $("resMsg").textContent = r === 1 ? "全對！太強了 🎉" : r >= .8 ? "很棒，差一點就全對！" : r >= .6 ? "不錯，看看錯的題目解析吧" : "再接再厲，先把解析看懂 💪";
    $("resPoints").textContent = quiz.mode === "practice" ? "再練一次不計分，錯的題目一樣會收進錯題本。"
      : rv ? "訂正成功 " + score + " 題，獲得 " + score + " 點。又錯的題目明天會再換題練一次。"
      : viewOnly ? "這一天答對 " + score + " 題，獲得 " + score + " 點。" : "今天獲得 " + score + " 點！答錯的題目明天起可以到錯題回收區練習。";
    $("resList").innerHTML = quiz.questions.map(function (q, i) {
      var x = quiz.results[q.id] || {}, ok = x.correct;
      return '<details class="res-item"' + (ok ? "" : " open") + "><summary>" +
        '<span class="mark ' + (ok ? "ok" : "bad") + '">' + (ok ? "○" : "✕") + "</span>" +
        '<span class="res-q">' + (i + 1) + ". " + qHtml(q.q) + "<br><small class=\"muted\">你的答案 (" + (x.choice || "-") + ") " +
        (x.choice ? esc(q.opts["ABCD".indexOf(x.choice)]) : "") + "　正解 (" + (x.ans || "?") + ") " + (x.ans ? esc(q.opts["ABCD".indexOf(x.ans)]) : "") + "</small></span>" +
        "</summary>" + (x.exp ? '<div class="res-exp">💬 ' + x.exp + "</div>" : "") + "</details>";
    }).join("");
    show("vResult");
  }

  // ---------- 啟動 ----------
  function boot() {
    if (DEMO) DemoApi.init();
    if (session) goHome(); else show("vLogin");
  }
  if (DEMO) {
    var n = 0;
    ["data.js", "variants.js"].forEach(function (f, i) {
      var s = document.createElement("script");
      s.src = f; s.async = false;
      s.onload = function () {
        if (++n < 2) return;
        if (!session) { $("inSid").value = "11001"; $("inName").value = "試看同學"; }
        boot();
      };
      document.head.appendChild(s);
    });
  } else boot();
})();
