(function () {
  "use strict";

  var CFG = window.APP_CONFIG || {};
  var LIVE = !!CFG.API_URL;
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

  // ---------- 後端：試看模式（瀏覽器內模擬，邏輯與 GAS 相同） ----------
  var DemoApi = (function () {
    var D, KEY = "grammarDemo_v1";
    function db() { return store(KEY) || { students: {} }; }
    function save(x) { store(KEY, x); }
    function today() { return $("demoDate").value; }
    function qmap() { var m = {}; D.questions.forEach(function (q) { m[q.id] = q; }); return m; }
    function umap() { var m = {}; D.units.forEach(function (u) { m[u.id] = u; }); return m; }
    function stu(x, sid, name) {
      if (!x.students[sid]) x.students[sid] = { name: name, points: 0, daily: {}, wrong: {} };
      return x.students[sid];
    }
    function pub(q) { return { id: q.id, q: q.q, opts: q.opts, src: q.src, unit: q.unit }; }
    function unitsFor(day, Q) {
      var ids = day.units && day.units.length ? day.units.slice() : [];
      if (!ids.length) day.qids.forEach(function (id) { var u = Q[id] && Q[id].unit; if (u && ids.indexOf(u) < 0) ids.push(u); });
      var U = umap();
      return ids.sort().map(function (id) { return U[id]; }).filter(Boolean);
    }
    function summary(s) {
      var t = today(), wd = parseDate(t).getDay();
      var days = D.schedule.map(function (d) {
        var a = s.daily[d.date] || {}, n = 0, sc = 0;
        d.qids.forEach(function (id) { if (a[id]) { n++; if (a[id].correct) sc++; } });
        var st = d.date > t ? "locked" : n >= d.qids.length ? "done" : n > 0 ? "partial" : "todo";
        return { date: d.date, title: d.title, total: d.qids.length, answered: n, score: sc, status: st };
      });
      var wrong = Object.keys(s.wrong).filter(function (k) { return s.wrong[k].status === "pending"; }).length;
      return {
        ok: true, today: t, points: s.points, wrongCount: wrong, days: days,
        reviewOpen: D.settings.reviewDays.indexOf(wd) >= 0, reviewDays: D.settings.reviewDays,
        doneDays: days.filter(function (d) { return d.status === "done"; }).length,
        student: { sid: session.sid, name: s.name }
      };
    }
    var api = {
      call: function (action, data) {
        return new Promise(function (resolve, reject) {
          setTimeout(function () {
            try { resolve(api[action](data || {})); } catch (e) { reject(e); }
          }, 120);
        });
      },
      login: function () { var x = db(); var s = stu(x, session.sid, session.name); save(x); return summary(s); },
      summary: function () { var x = db(); return summary(stu(x, session.sid, session.name)); },
      getDay: function (p) {
        var day = D.schedule.filter(function (d) { return d.date === p.date; })[0];
        if (!day) throw new Error("這一天沒有排練習");
        if (day.date > today()) throw new Error("這一天還沒開放");
        var Q = qmap(), x = db(), s = stu(x, session.sid, session.name), a = s.daily[day.date] || {};
        var answered = {};
        day.qids.forEach(function (id) { if (a[id]) answered[id] = { choice: a[id].choice, correct: a[id].correct, ans: Q[id].ans, exp: Q[id].exp }; });
        return { ok: true, date: day.date, title: day.title, units: unitsFor(day, Q),
          questions: day.qids.map(function (id) { return pub(Q[id]); }), answered: answered };
      },
      getReview: function () {
        if (!summary(stu(db(), session.sid, session.name)).reviewOpen) throw new Error("錯題回收在週末開放");
        var Q = qmap(), x = db(), s = stu(x, session.sid, session.name);
        var ids = Object.keys(s.wrong).filter(function (k) { return s.wrong[k].status === "pending"; })
          .sort(function (a, b) { return s.wrong[a].t - s.wrong[b].t; }).slice(0, D.settings.reviewCount);
        return { ok: true, title: "週末錯題回收", units: unitsFor({ units: [], qids: ids }, Q),
          questions: ids.map(function (id) { return pub(Q[id]); }), answered: {} };
      },
      answer: function (p) {
        var Q = qmap(), q = Q[p.qid]; if (!q) throw new Error("找不到題目");
        var x = db(), s = stu(x, session.sid, session.name), correct = p.choice === q.ans, gained = 0;
        if (p.mode === "daily") {
          var a = s.daily[p.date] = s.daily[p.date] || {};
          if (a[p.qid]) return { ok: true, correct: a[p.qid].correct, ans: q.ans, exp: q.exp, gained: 0, points: s.points, repeat: true };
          a[p.qid] = { choice: p.choice, correct: correct };
          if (correct) gained = 1;
        }
        if (p.mode === "review") {
          var w = s.wrong[p.qid];
          if (w && w.status === "pending") { if (correct) { w.status = "fixed"; gained = 1; } else w.count++; }
        } else if (!correct) {
          var ww = s.wrong[p.qid];
          if (!ww || ww.status !== "pending") s.wrong[p.qid] = { count: 1, status: "pending", t: Date.now() };
          else ww.count++;
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

  var Api = LIVE ? LiveApi : DemoApi;

  // ---------- 狀態 ----------
  var session = store("grammarSession");
  var home = null;   // 最近一次 summary
  var quiz = null;   // { mode, date, title, units, questions, idx, results }

  function show(v) {
    ["vLogin", "vHome", "vNotes", "vQuiz", "vResult"].forEach(function (id) { $(id).hidden = id !== v; });
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
      store("grammarSession", session);
      renderHome(res);
    }).catch(function (err) { loading(false); session = null; $("loginErr").textContent = err.message; });
  });
  $("logoutBtn").addEventListener("click", function () { session = null; store("grammarSession", null); $("inSid").value = ""; $("inName").value = ""; show("vLogin"); });

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
    $("statWrong").textContent = res.wrongCount;

    var t = res.today, todayDay = res.days.filter(function (d) { return d.date === t; })[0];
    var card = $("todayCard"), h = "";
    card.className = "card today-card";
    if (todayDay) {
      h += '<div class="eyebrow">今天 ' + niceDate(t) + " · 第 " + (res.days.indexOf(todayDay) + 1) + " 天</div>";
      h += "<h2>" + esc(todayDay.title) + "</h2>";
      h += statusBadge(todayDay);
      h += '<div class="btn-row">' + dayButtons(todayDay) + "</div>";
    } else if (res.reviewOpen) {
      card.className += " review";
      h += '<div class="eyebrow" style="color:var(--review)">今天 ' + niceDate(t) + " · 週末</div>";
      h += "<h2>錯題回收時間 ♻️</h2>";
      if (res.wrongCount > 0) {
        h += '<span class="status review">錯題本還有 ' + res.wrongCount + " 題</span>";
        h += '<p class="muted">把這週答錯的題目再做一次，答對就從錯題本移除，並得 1 點。</p>';
        h += '<button class="btn review big" data-review="1">開始錯題回收（最多 ' + Math.min(10, res.wrongCount) + " 題）</button>";
      } else {
        h += '<p class="muted">錯題本是空的，太厲害了！可以從下方月曆補做還沒完成的練習。</p>';
      }
    } else {
      var first = res.days[0], last = res.days[res.days.length - 1];
      h += '<div class="eyebrow">今天 ' + niceDate(t) + "</div>";
      if (first && t < first.date) h += "<h2>練習從 " + niceDate(first.date) + " 開始</h2><p class=\"muted\">準備好了嗎？到時候每天 10 題，答對一題得 1 點！</p>";
      else if (last && t > last.date) h += "<h2>本月練習已全部結束 🎉</h2><p class=\"muted\">可以從月曆回顧每天的題目。</p>";
      else h += "<h2>今天沒有排練習</h2><p class=\"muted\">可以從下方月曆補做還沒完成的練習。</p>";
    }
    if (!res.reviewOpen && res.wrongCount > 0 && todayDay) {
      h += '<p class="muted" style="margin:12px 0 0;font-size:14px">📕 錯題本有 ' + res.wrongCount + " 題，週末開放錯題回收。</p>";
    }
    card.innerHTML = h;

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
    if (d.status === "partial") return '<span class="status partial">做到第 ' + d.answered + " 題 · 先讀文法再繼續</span>";
    return '<span class="status todo">尚未開始 · 10 題</span>';
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
      var ds = fmtDate(new Date(y, m, dd)), wd = new Date(y, m, dd).getDay(), d = byDate[ds];
      var cls = "day", sub = "", attr = "", isToday = ds === res.today;
      if (d) {
        cls += " " + d.status;
        sub = d.status === "done" ? d.score + "/" + d.total : d.status === "partial" ? d.answered + "/" + d.total : d.status === "locked" ? "" : "待做";
        if (d.status !== "locked") attr = d.status === "done" ? ' data-result="' + ds + '"' : ' data-day="' + ds + '"';
      } else if (res.reviewDays.indexOf(wd) >= 0 && res.days.length && ds > res.days[0].date) {
        cls += " review"; sub = "錯題";
        if (isToday && res.reviewOpen && res.wrongCount > 0) attr = ' data-review="1"';
        if (ds > res.today) cls += " locked";
      } else cls += " locked";
      if (isToday) cls += " today";
      h += attr ? '<button class="' + cls + '"' + attr + ">" + dd + "<small>" + sub + "</small></button>"
                : '<div class="' + cls + '">' + dd + "<small>" + sub + "</small></div>";
    }
    $("cal").innerHTML = h;
  }

  // 首頁按鈕（事件委派）
  document.addEventListener("click", function (e) {
    var b = e.target.closest("[data-day],[data-redo],[data-result],[data-review],[data-go]");
    if (!b) return;
    if (b.dataset.go === "home") return goHome();
    if (b.dataset.day) return openDay(b.dataset.day, "daily");
    if (b.dataset.redo) return openDay(b.dataset.redo, "practice");
    if (b.dataset.result) return openResult(b.dataset.result);
    if (b.dataset.review) return openReview();
  });

  // ---------- 開始一天的練習 ----------
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
      if (!res.questions.length) { alert("錯題本是空的！"); return goHome(); }
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
    var q = quiz.questions[quiz.idx], n = quiz.questions.length;
    $("quizCount").textContent = (quiz.mode === "review" ? "錯題 " : "第 ") + (quiz.idx + 1) + " / " + n + (quiz.mode === "review" ? "" : " 題");
    $("progressBar").style.width = (quiz.idx / n * 100) + "%";
    $("progressBar").style.background = quiz.mode === "review" ? "var(--review)" : "";
    $("qSrc").textContent = q.id.slice(0, 2) + " · " + q.src;
    $("qText").innerHTML = qHtml(q.q);
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
    var q = quiz.questions[quiz.idx], choice = b.dataset.c;
    var all = $("qOpts").querySelectorAll(".opt");
    all.forEach(function (x) { x.disabled = true; });
    b.style.borderColor = "var(--primary)";
    Api.call("answer", { mode: quiz.mode, date: quiz.date, qid: q.id, choice: choice }).then(function (res) {
      quiz.results[q.id] = { choice: choice, correct: res.correct, ans: res.ans, exp: res.exp };
      b.style.borderColor = "";
      all.forEach(function (x) {
        if (x.dataset.c === res.ans) x.classList.add("correct");
        else if (x.dataset.c === choice) x.classList.add("wrong");
      });
      var fb = $("qFeedback");
      fb.className = "feedback " + (res.correct ? "ok" : "bad");
      var head = res.correct ? (quiz.mode === "review" ? "訂正成功！" : "答對了！") : "答錯了，正確答案是 (" + res.ans + ")";
      if (res.gained) head += "　+" + res.gained + " 點";
      if (!res.correct && quiz.mode !== "review") head += '<div style="font-size:13px;font-weight:500;color:var(--muted)">已收進錯題本，週末再來挑戰</div>';
      if (!res.correct && quiz.mode === "review") head += '<div style="font-size:13px;font-weight:500;color:var(--muted)">這題會留在錯題本，下次再試一次</div>';
      fb.innerHTML = '<div class="fb-head">' + head + '</div><div class="exp">💬 ' + res.exp + "</div>";
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
      alert("送出失敗，請再按一次：" + err.message);
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
    var total = quiz.questions.length, score = 0;
    quiz.questions.forEach(function (q) { var r = quiz.results[q.id]; if (r && r.correct) score++; });
    $("resEyebrow").textContent = quiz.mode === "review" ? "錯題回收結果" : niceDate(quiz.date) + " · " + quiz.title;
    $("resScore").textContent = score;
    $("resTotal").textContent = total;
    document.querySelector(".score-ring").style.setProperty("--pct", (score / total * 100) + "%");
    var r = score / total;
    $("resMsg").textContent = r === 1 ? "全對！太強了 🎉" : r >= .8 ? "很棒，差一點就全對！" : r >= .6 ? "不錯，看看錯的題目解析吧" : "再接再厲，先把解析看懂 💪";
    $("resPoints").textContent = quiz.mode === "practice" ? "再練一次不計分，錯的題目一樣會收進錯題本。"
      : quiz.mode === "review" ? "訂正成功 " + score + " 題，獲得 " + score + " 點。"
      : viewOnly ? "這一天答對 " + score + " 題，獲得 " + score + " 點。" : "今天獲得 " + score + " 點！答錯的題目已收進錯題本。";
    $("resList").innerHTML = quiz.questions.map(function (q, i) {
      var x = quiz.results[q.id] || {};
      var ok = x.correct;
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
    if (!LIVE) DemoApi.init();
    if (session) { goHome(); } else show("vLogin");
  }
  if (!LIVE) {
    var s = document.createElement("script");
    s.src = "data.js";
    s.onload = function () {
      if (window.GRAMMAR_DATA.settings.title) $("appTitle").textContent = window.GRAMMAR_DATA.settings.title;
      if (!session) { $("inSid").value = "11001"; $("inName").value = "試看同學"; }
      boot();
    };
    document.head.appendChild(s);
  } else boot();
})();
