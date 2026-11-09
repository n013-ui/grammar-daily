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
  // 克漏字文章：開頭 [[n]] 表示這一題是第 n 格；{k} 是各個空格
  function passageHtml(p) {
    if (!p) return "";
    var m = p.match(/^\[\[(\d+)\]\]/), cur = m ? m[1] : "";
    var body = esc(p.replace(/^\[\[\d+\]\]/, "")).replace(/\{(\d+)\}/g, function (_, k) {
      return '<span class="gap' + (k === cur ? " cur" : "") + '">' + (k === cur ? "&nbsp;" + k + "&nbsp;" : k) + "</span>";
    }).replace(/\n/g, "<br>");
    return '<div class="passage">' + body + "</div>";
  }
  // 答錯時：考點、你可能是…、盲點提醒
  function insightHtml(r) {
    if (!r || (!r.kpName && !r.trap)) return "";
    var h = '<div class="insight">';
    if (r.kpName) h += '<div>🎯 <b>考點</b>：' + esc(r.kpName) + "</div>";
    if (!r.correct && r.trap) h += '<div class="maybe">🤔 <b>你可能是</b>：' + esc(r.trap) + "</div>";
    if (!r.correct && r.kpTip) h += '<div class="tipline">📌 ' + esc(r.kpTip) + "</div>";
    return h + "</div>";
  }
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

  // ---------- 後端：試看模式（瀏覽器裡直接執行 gas/Code.gs，規則與正式版相同；只在本機用） ----------
  var DemoApi = {
    call: function (action, data) {
      data = data || {};
      if (session) { data.sid = session.sid; data.name = session.name; }
      return new Promise(function (resolve, reject) {
        setTimeout(function () {
          try { var r = window.DemoGas.call(action, data); if (!r.ok) throw new Error(r.error || "發生錯誤"); resolve(r); } catch (e) { reject(e); }
        }, 80);
      });
    },
    init: function () {
      var d = $("demoDate"), first = "2026-11-09", last = "2026-11-30";
      var t = fmtDate(new Date());
      d.min = "2026-11-01"; d.max = last;
      d.value = store("grammarDemoDate") || (t < first ? first : t > last ? last : t);
      d.addEventListener("change", function () { store("grammarDemoDate", d.value); if (session) goHome(); });
      $("demoReset").addEventListener("click", function () {
        if (!confirm("清除這台電腦上所有試看模式的作答紀錄？")) return;
        window.DemoGas.reset(); if (session) goHome();
      });
      $("demoBar").hidden = false;
      return window.DemoGas.load(function () { return d.value; });
    }
  };

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
    var sid = $("inSid").value.trim();
    if (!sid) return;
    session = { sid: sid, name: sid };
    $("loginErr").textContent = "";
    loading(true);
    Api.call("login").then(function (res) {
      loading(false);
      session.sid = res.student.sid; session.name = res.student.name;
      store(SKEY, session);
      renderHome(res);
    }).catch(function (err) { loading(false); session = null; $("loginErr").textContent = err.message; });
  });
  $("logoutBtn").addEventListener("click", function () { session = null; store(SKEY, null); $("inSid").value = ""; show("vLogin"); });

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
      var nextDay = res.days.filter(function (d) { return d.date > t && d.status !== "done"; })[0];
      if (todayDay.status === "done" && nextDay) h += '<p class="muted" style="margin:10px 0 0">還想練？可以先做 <button class="link" data-day="' + nextDay.date + '">' + niceDate(nextDay.date) + " 的題目 →</button></p>";
    } else {
      var first = res.days[0], last = res.days[res.days.length - 1];
      h += '<div class="eyebrow">今天 ' + niceDate(t) + "</div>";
      if (first && t < first.date) h += "<h2>練習從 " + niceDate(first.date) + " 開始</h2><p class=\"muted\">到時候每天 20 題，答對一題得 1 點！</p>";
      else if (last && t > last.date) h += "<h2>本月練習已全部結束 🎉</h2><p class=\"muted\">可以到錯題回收區把錯題清光，或從月曆回顧題目。</p>";
      else if (wd === 0 || wd === 6) h += "<h2>週末沒有排新進度 ☕</h2><p class=\"muted\">可以到下面的錯題回收區練習、補做還沒完成的天數，或從月曆先做下週的題目。</p>";
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

    // 考前特訓
    var dc = res.drill || { L3: 0, L4: 0, ALL: 0 };
    var g = '<h3>🎯 考前特訓</h3><p class="muted" style="margin:4px 0 0">把這一課<b>曾經錯過的題目</b>（包括已經訂正的）再練一遍，加上同觀念的其他題目。錯越多的觀念出越多題。不計分，適合段考前衝刺。</p>';
    var un = dc.unlock || {}, nd = dc.need || {};
    g += '<div class="btn-row drill-btns">' + [["L3", "第三課"], ["L4", "第四課"], ["ALL", "兩課一起"]].map(function (x) {
      var k = x[0], open = un[k] && dc[k] > 0, sub;
      if (!un[k]) sub = k === "ALL" ? "🔒 兩課都完成後解鎖" : "🔒 完成 " + (nd[k] ? nd[k][0] + " / " + nd[k][1] : "") + " 天後解鎖";
      else sub = dc[k] ? "錯過 " + dc[k] + " 題" : "沒有錯題 👍";
      return '<button class="btn drill big' + (un[k] ? "" : " locked") + '" data-drill="' + k + '"' + (open ? "" : " disabled") + ">" + (un[k] ? "" : "🔒 ") + x[1] +
        "<small>" + sub + "</small></button>";
    }).join("") + "</div>";
    g += '<p class="muted" style="margin:8px 0 0;font-size:13px">要先把這一課的每日練習全部做完，才能解鎖那一課的考前特訓。</p>';
    $("drillCard").innerHTML = g;

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
    return '<span class="status todo">' + (d.status === "ahead" ? "可以先做 · " : "尚未開始 · ") + d.total + " 題</span>";
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
        sub = d.status === "done" ? d.score + "/" + d.total : d.status === "partial" ? d.answered + "/" + d.total : d.status === "ahead" ? "先做" : "待做";
        attr = d.status === "done" ? ' data-result="' + ds + '"' : ' data-day="' + ds + '"';
      } else cls += " locked";
      if (ds === res.today) cls += " today";
      h += attr ? '<button class="' + cls + '"' + attr + ">" + dd + "<small>" + sub + "</small></button>"
                : '<div class="' + cls + '">' + dd + "<small>" + sub + "</small></div>";
    }
    $("cal").innerHTML = h;
  }

  // 按鈕（事件委派）
  document.addEventListener("click", function (e) {
    var b = e.target.closest("[data-day],[data-redo],[data-result],[data-review],[data-book],[data-go],[data-drill]");
    if (!b) return;
    if (b.dataset.go === "home") return goHome();
    if (b.dataset.day) return openDay(b.dataset.day, "daily");
    if (b.dataset.redo) return openDay(b.dataset.redo, "practice");
    if (b.dataset.result) return openResult(b.dataset.result);
    if (b.dataset.review) return openReview();
    if (b.dataset.book) return openBook();
    if (b.dataset.drill && !b.disabled) return openDrill(b.dataset.drill);
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
        return '<div class="book-item' + (w.status === "已訂正" ? " fixed" : "") + '">' + (w.kpName ? '<div class="kp-tag">🎯 ' + esc(w.kpName) + "</div>" : "") + '<div class="bq">' + qHtml(w.q) + '</div><div class="bmeta">' + st + again +
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
  function openDrill(lesson) {
    loading(true);
    Api.call("getDrill", { lesson: lesson }).then(function (res) {
      loading(false);
      quiz = { mode: "drill", date: home.today, title: res.title, units: res.units, questions: res.questions, results: {}, idx: 0 };
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
    if (quiz.date > home.today && quiz.mode === "daily") $("notesEyebrow").textContent = niceDate(quiz.date) + " 的進度 · 提前練習（點數一樣算）";
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
    var q = quiz.questions[quiz.idx], n = quiz.questions.length, rv = quiz.mode === "review", dr = quiz.mode === "drill";
    $("quizCount").textContent = (rv ? "錯題 " : "第 ") + (quiz.idx + 1) + " / " + n + (rv ? "" : " 題");
    $("progressBar").style.width = (quiz.idx / n * 100) + "%";
    $("progressBar").style.background = rv ? "var(--review)" : dr ? "#c2410c" : "";
    var meta = "";
    if (rv) {
      meta = '<div class="qbadges"><span class="badge">🔁 第 ' + (q.attempts + 1) + " 次回收練習</span>" +
        (q.src === "變化題" ? '<span class="badge">換句話考同觀念</span>' : '<span class="badge">原題再挑戰（選項已換位置）</span>') +
        '<span class="badge bad">已錯 ' + q.wrongCount + " 次</span>" +
        (q.hist && q.hist.slice(-1) === "✗" ? '<span class="badge bad">上次又錯了</span>' : "") + "</div>";
    }
    if (dr) meta = '<div class="qbadges">' + (q.drillType === "wrong" ? '<span class="badge bad">🎯 你錯過這題 ' + q.wrongCount + " 次</span>"
      : '<span class="badge">同觀念練習</span>') + "</div>";
    $("qSrc").textContent = (rv ? "錯題回收 · " : dr ? "考前特訓 · " : "") + (q.kpName || q.src || "");
    $("qText").innerHTML = meta + passageHtml(q.passage) + qHtml(q.q);
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
    var q = quiz.questions[quiz.idx], choice = b.dataset.c, rv = quiz.mode === "review", dr = quiz.mode === "drill";
    var all = $("qOpts").querySelectorAll(".opt");
    all.forEach(function (x) { x.disabled = true; });
    b.style.borderColor = "var(--primary)";
    var payload = { mode: quiz.mode, date: quiz.date, qid: q.id, choice: choice };
    if (rv) payload.vk = q.vk;
    if (rv || dr) payload.choiceText = q.opts["ABCD".indexOf(choice)];
    Api.call("answer", payload).then(function (res) {
      var ansL = res.ansText != null ? "ABCD"[q.opts.indexOf(res.ansText)] : res.ans;
      quiz.results[q.id] = { choice: choice, correct: res.correct, ans: ansL, exp: res.exp, kpName: res.kpName, kpTip: res.kpTip, trap: res.trap };
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
      if (dr) sub = res.correct ? "考前特訓不計分，觀念有記住就好 👍" : "已記到錯題本，考前再看一次詳解";
      if (!res.correct && rv) sub = "這題已錯 " + res.wrongCount + " 次，會留在錯題本，明天再換題練一次";
      if (res.correct && rv) sub = "已從錯題本移除。回收紀錄：" + res.hist;
      fb.innerHTML = '<div class="fb-head">' + head + (sub ? '<div style="font-size:13px;font-weight:500;color:var(--muted)">' + sub + "</div>" : "") +
        "</div>" + insightHtml(res) + '<div class="exp">💬 <b>詳解</b>　' + res.exp + "</div>";
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
    var total = quiz.questions.length, score = 0, rv = quiz.mode === "review", dr = quiz.mode === "drill";
    quiz.questions.forEach(function (q) { var r = quiz.results[q.id]; if (r && r.correct) score++; });
    $("resEyebrow").textContent = rv ? "錯題回收結果" : dr ? quiz.title + " 結果" : niceDate(quiz.date) + " · " + quiz.title;
    $("resScore").textContent = score;
    $("resTotal").textContent = total;
    document.querySelector(".score-ring").style.setProperty("--pct", (score / total * 100) + "%");
    var r = score / total;
    $("resMsg").textContent = r === 1 ? "全對！太強了 🎉" : r >= .8 ? "很棒，差一點就全對！" : r >= .6 ? "不錯，看看錯的題目解析吧" : "再接再厲，先把解析看懂 💪";
    $("resPoints").textContent = quiz.mode === "practice" ? "再練一次不計分，錯的題目一樣會收進錯題本。"
      : dr ? "考前特訓不計分。錯的題目已記到錯題本，把下面的詳解看懂再去考試！"
      : rv ? "訂正成功 " + score + " 題，獲得 " + score + " 點。又錯的題目明天會再換題練一次。"
      : viewOnly ? "這一天答對 " + score + " 題，獲得 " + score + " 點。" : "今天獲得 " + score + " 點！答錯的題目明天起可以到錯題回收區練習。";
    $("resList").innerHTML = quiz.questions.map(function (q, i) {
      var x = quiz.results[q.id] || {}, ok = x.correct;
      return '<details class="res-item"' + (ok ? "" : " open") + "><summary>" +
        '<span class="mark ' + (ok ? "ok" : "bad") + '">' + (ok ? "○" : "✕") + "</span>" +
        '<span class="res-q">' + (i + 1) + ". " + (q.passage ? "［克漏字 第 " + (q.passage.match(/^\[\[(\d+)\]\]/) || [0, "?"])[1] + " 格］" : qHtml(q.q)) + "<br><small class=\"muted\">你的答案 (" + (x.choice || "-") + ") " +
        (x.choice ? esc(q.opts["ABCD".indexOf(x.choice)]) : "") + "　正解 (" + (x.ans || "?") + ") " + (x.ans ? esc(q.opts["ABCD".indexOf(x.ans)]) : "") + "</small></span>" +
        "</summary>" + insightHtml(x) + (x.exp ? '<div class="res-exp">💬 ' + x.exp + "</div>" : "") + "</details>";
    }).join("");
    show("vResult");
  }

  // ---------- 啟動 ----------
  function boot() {
    if (session) goHome(); else show("vLogin");
  }
  if (DEMO) {
    var sc = document.createElement("script");
    sc.src = "demo-gas.js";
    sc.onload = function () {
      DemoApi.init().then(function () {
        if (!session) $("inSid").value = "queena";
        boot();
      }).catch(function (e) { alert("試看模式載入失敗：" + e.message); });
    };
    document.head.appendChild(sc);
  } else boot();
})();
