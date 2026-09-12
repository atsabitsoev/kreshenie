/* ============================================================
   Курс подготовки ко крещению — интерфейс ученика
   ============================================================ */
(function () {
  'use strict';

  var C = window.COURSE;
  var DAYS = {};
  C.modules.forEach(function (m) { m.days.forEach(function (d) { DAYS[d.id] = d; }); });

  var S = {
    token: '',
    student: null,
    submitted: {},          // dayId -> {answers, submittedAt}
    values: {},             // dayId -> {key: value}
    currentId: null,
    dirty: false,
    saveTimer: null,
    syncTimer: null,
    view: 'day'             // 'day' | 'appendix'
  };

  /* ───────────────────────── Утилиты ───────────────────────── */

  var $ = function (sel, root) { return (root || document).querySelector(sel); };
  function el(tag, cls, html) {
    var n = document.createElement(tag);
    if (cls) n.className = cls;
    if (html != null) n.innerHTML = html;
    return n;
  }
  function esc(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }
  function lsKey() { return 'krsh:' + S.token; }
  function loadLocal() {
    try { return JSON.parse(localStorage.getItem(lsKey()) || '{}'); } catch (e) { return {}; }
  }
  function saveLocal() {
    try { localStorage.setItem(lsKey(), JSON.stringify(S.values)); } catch (e) {}
  }
  function idx(dayId) { return C.order.indexOf(dayId); }
  function isDone(dayId) { return !!S.submitted[dayId]; }
  function isUnlocked(dayId) {
    var i = idx(dayId);
    if (i <= 0) return true;
    return isDone(C.order[i - 1]) || isDone(dayId);
  }
  function doneCount() { return C.order.filter(isDone).length; }
  function firstOpenDay() {
    for (var i = 0; i < C.order.length; i++) if (!isDone(C.order[i])) return C.order[i];
    return C.order[C.order.length - 1];
  }
  function fmtDate(iso) {
    if (!iso) return '';
    var d = new Date(iso);
    if (isNaN(d)) return '';
    return d.toLocaleDateString('ru-RU', { day: 'numeric', month: 'long' }) + ', ' +
           d.toLocaleTimeString('ru-RU', { hour: '2-digit', minute: '2-digit' });
  }
  var toastTimer;
  function toast(msg) {
    var t = $('#toast');
    t.textContent = msg; t.classList.add('show');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { t.classList.remove('show'); }, 2600);
  }

  /* ───────────────────────── Тема ───────────────────────── */

  function initTheme() {
    var saved = null;
    try { saved = localStorage.getItem('krsh:theme'); } catch (e) {}
    if (saved) document.documentElement.setAttribute('data-theme', saved);
    paintThemeBtn();
  }
  function paintThemeBtn() {
    var b = $('#themeBtn');
    if (!b) return;
    var dark = document.documentElement.getAttribute('data-theme') === 'dark' ||
      (!document.documentElement.getAttribute('data-theme') &&
        window.matchMedia('(prefers-color-scheme: dark)').matches);
    b.innerHTML = dark ? ICON.sun : ICON.moon;
    b.setAttribute('aria-label', dark ? 'Светлая тема' : 'Тёмная тема');
  }
  function toggleTheme() {
    var cur = document.documentElement.getAttribute('data-theme');
    var dark = cur === 'dark' || (!cur && window.matchMedia('(prefers-color-scheme: dark)').matches);
    var next = dark ? 'light' : 'dark';
    document.documentElement.setAttribute('data-theme', next);
    try { localStorage.setItem('krsh:theme', next); } catch (e) {}
    paintThemeBtn();
  }

  /* ───────────────────────── Экраны ───────────────────────── */

  function org() { return (window.CONFIG && window.CONFIG.ORG) || ''; }

  function screenCentered(html) {
    document.body.innerHTML = '<div class="centered"><div class="card">' + html + '</div></div>' +
      '<div class="toast" id="toast"></div>';
  }

  function screenWelcome(errMsg) {
    screenCentered(
      '<div class="mark-lg">' + ICON.book + '</div>' +
      (org() ? '<div class="org-top">' + esc(org()) + '</div>' : '') +
      '<h1>' + esc(C.title) + '</h1>' +
      '<p>Чтобы начать, откройте персональную ссылку, которую дал вам наставник, или введите её код ниже.</p>' +
      (errMsg ? '<p style="color:var(--danger);margin-top:14px">' + esc(errMsg) + '</p>' : '') +
      '<div class="form">' +
      '  <input class="input" id="tokenInput" placeholder="Код из ссылки" autocomplete="off" spellcheck="false">' +
      '  <button class="btn btn-lg btn-block" id="goBtn" style="margin-top:12px">Продолжить</button>' +
      '</div>'
    );
    var go = function () {
      var v = $('#tokenInput').value.trim();
      var m = v.match(/[?&]s=([^&\s]+)/);
      if (m) v = decodeURIComponent(m[1]);
      if (!v) return;
      location.search = '?s=' + encodeURIComponent(v);
    };
    $('#goBtn').onclick = go;
    $('#tokenInput').onkeydown = function (e) { if (e.key === 'Enter') go(); };
  }

  function screenLoading() {
    screenCentered('<div class="spinner"></div><p style="margin-top:18px">Загружаем ваш курс…</p>');
  }

  function screenError(title, text, retry) {
    screenCentered(
      '<div class="mark-lg">' + ICON.info + '</div>' +
      '<h1>' + esc(title) + '</h1><p>' + text + '</p>' +
      (retry ? '<button class="btn btn-lg btn-block" id="retryBtn" style="margin-top:22px">Попробовать снова</button>' : '')
    );
    if (retry) $('#retryBtn').onclick = function () { location.reload(); };
  }

  /* ───────────────────────── Каркас приложения ───────────────────────── */

  function shell() {
    document.body.innerHTML =
      '<header class="topbar with-sidebar">' +
      '  <button class="icon-btn menu-btn" id="menuBtn" aria-label="Меню">' + ICON.menu + '</button>' +
      '  <div class="topbar-title"><strong id="tbTitle"></strong><span id="tbSub"></span></div>' +
      '  <div class="progress-pill" id="progPill"></div>' +
      '  <button class="icon-btn" id="themeBtn"></button>' +
      '</header>' +
      '<div class="scrim" id="scrim"></div>' +
      '<aside class="sidebar" id="sidebar">' +
      '  <div class="sidebar-head"><strong>' + esc(C.title) + '</strong>' +
      '    <button class="icon-btn" id="closeNav" aria-label="Закрыть">' + ICON.close + '</button></div>' +
      '  <div class="sidebar-scroll" id="navList"></div>' +
      (org() ? '<div class="sidebar-foot">' + esc(org()) + '</div>' : '') +
      '</aside>' +
      '<div class="layout"><main class="main"><div class="wrap" id="view"></div></main></div>' +
      '<div class="toast" id="toast"></div>';

    $('#menuBtn').onclick = openNav;
    $('#closeNav').onclick = closeNav;
    $('#scrim').onclick = closeNav;
    $('#themeBtn').onclick = toggleTheme;
    paintThemeBtn();
    document.addEventListener('keydown', function (e) { if (e.key === 'Escape') closeNav(); });
  }

  function openNav() { $('#sidebar').classList.add('open'); $('#scrim').classList.add('show'); }
  function closeNav() { $('#sidebar').classList.remove('open'); $('#scrim').classList.remove('show'); }

  function paintTop() {
    var n = doneCount(), total = C.totalDays;
    $('#tbTitle').textContent = S.student ? S.student.name : C.title;
    $('#tbSub').textContent = n >= total ? 'Курс пройден' : 'День ' + (idx(S.currentId) + 1) + ' из ' + total;
    var r = 9, circ = 2 * Math.PI * r, pct = total ? n / total : 0;
    $('#progPill').innerHTML =
      '<svg class="ring" viewBox="0 0 22 22"><circle class="track" cx="11" cy="11" r="' + r + '"/>' +
      '<circle class="bar" cx="11" cy="11" r="' + r + '" stroke-dasharray="' +
      (circ * pct).toFixed(1) + ' ' + circ.toFixed(1) + '"/></svg>' +
      '<span>' + n + '/' + total + '</span>';
  }

  function paintNav() {
    var wrap = el('div');
    C.modules.forEach(function (m) {
      var box = el('div', 'nav-module');
      box.appendChild(el('div', 'nav-module-head',
        '<b>' + m.num + '. ' + esc(m.title) + '</b>'));
      m.days.forEach(function (d) {
        var done = isDone(d.id), open = isUnlocked(d.id), cur = d.id === S.currentId && S.view === 'day';
        var b = el('button', 'nav-day' + (done ? ' done' : '') + (cur ? ' current' : ''));
        b.disabled = !open;
        b.innerHTML =
          '<span class="mark">' + (done ? ICON.check : d.dayInModule) + '</span>' +
          '<span class="label">' + esc(d.title) + '</span>' +
          (open ? '' : '<span class="lock">' + ICON.lock + '</span>');
        b.onclick = function () { goDay(d.id); closeNav(); };
        box.appendChild(b);
      });
      wrap.appendChild(box);
    });

    var ax = el('div', 'nav-module');
    ax.appendChild(el('div', 'nav-module-head', '<b>Приложение</b>'));
    var ab = el('button', 'nav-day' + (S.view === 'appendix' ? ' current' : ''));
    ab.innerHTML = '<span class="mark">' + ICON.book + '</span><span class="label">Материалы для благовестия</span>';
    ab.onclick = function () { goAppendix(); closeNav(); };
    ax.appendChild(ab);
    wrap.appendChild(ax);

    var list = $('#navList');
    list.innerHTML = '';
    list.appendChild(wrap);
  }

  /* ───────────────────────── Значения полей ───────────────────────── */

  function valuesFor(dayId) {
    if (!S.values[dayId]) S.values[dayId] = {};
    return S.values[dayId];
  }
  function setVal(dayId, key, v) {
    valuesFor(dayId)[key] = v;
    S.dirty = true;
    saveLocal();
    scheduleSync(dayId);
    flagStatus('saving');
  }

  function scheduleSync(dayId) {
    clearTimeout(S.syncTimer);
    S.syncTimer = setTimeout(function () {
      API.post({ action: 'draft', token: S.token, dayId: dayId, answers: valuesFor(dayId) })
        .then(function () { flagStatus('saved'); })
        .catch(function () { flagStatus('local'); });
    }, 4000);
  }

  function flagStatus(kind) {
    var n = $('#saveStatus');
    if (!n) return;
    n.className = 'status' + (kind === 'error' ? ' error' : kind === 'saving' ? ' saving' : '');
    n.textContent = kind === 'saving' ? 'Сохраняем…'
      : kind === 'saved' ? 'Черновик сохранён'
      : kind === 'local' ? 'Сохранено на устройстве'
      : '';
  }

  /* ───────────────────────── Рендер дня ───────────────────────── */

  function goDay(dayId) {
    S.currentId = dayId; S.view = 'day';
    history.replaceState(null, '', '?s=' + encodeURIComponent(S.token) + '#' + dayId);
    renderDay();
    paintNav(); paintTop();
    window.scrollTo({ top: 0, behavior: 'instant' in window ? 'instant' : 'auto' });
  }

  function renderDay() {
    var d = DAYS[S.currentId];
    var done = isDone(d.id);
    var vals = done ? (S.submitted[d.id].answers || {}) : valuesFor(d.id);
    var v = $('#view');
    v.innerHTML = '';

    // Шапка
    var head = el('div', 'day-head');
    head.innerHTML =
      '<div class="eyebrow">' + esc(d.moduleTitle) + '<span class="dot"></span>День ' +
      d.dayInModule + ' из ' + d.daysInModule + '</div>' +
      '<h1>' + esc(d.title) + '</h1>' +
      (d.subtitle ? '<div class="sub">' + esc(d.subtitle) + '</div>' : '');
    v.appendChild(head);

    if (doneCount() >= C.totalDays) {
      var fin = el('div', 'finish');
      fin.innerHTML = '<div class="ic">' + ICON.check + '</div>' +
        '<h2>Курс пройден</h2><p>Все ' + C.totalDays + ' дней позади. ' +
        'Ответы сохранены — вы можете перечитать любой день, а наставник видит весь ваш путь.</p>';
      v.appendChild(fin);
    }

    if (done && doneCount() < C.totalDays) {
      var db = el('div', 'done-banner');
      db.innerHTML = ICON.check + '<div><b>День пройден</b>Отправлено ' +
        esc(fmtDate(S.submitted[d.id].submittedAt)) + '. Ответы сохранены у наставника.</div>';
      v.appendChild(db);
    }

    if (d.intro && d.intro.length) {
      var intro = el('div', 'intro');
      intro.innerHTML = d.intro.map(function (p) { return '<p>' + p + '</p>'; }).join('');
      v.appendChild(intro);
    }

    if (d.verse) {
      var vs = el('div', 'verse');
      vs.innerHTML = '<div class="verse-text">' + d.verse.text + '</div>' +
                     '<div class="verse-ref">' + esc(d.verse.ref) + '</div>';
      v.appendChild(vs);
    }

    if (d.read) {
      var rd = el('div', 'read');
      rd.innerHTML = ICON.book + '<div class="t">' + esc(d.read.label) +
        '<b>' + esc(d.read.ref) + '</b></div>';
      v.appendChild(rd);
    }

    (d.blocks || []).forEach(function (b) {
      var node = renderBlock(b, d, vals, done);
      if (node) v.appendChild(node);
    });

    if (d.notes && d.notes.length) {
      var nt = el('div', 'notes');
      nt.innerHTML = '<h3>Пояснения</h3>' + d.notes.map(function (n) {
        return '<div class="note"><b>' + esc(n.term) + '</b> — ' + n.text +
          (n.src ? '<span class="src">' + esc(n.src) + '</span>' : '') + '</div>';
      }).join('');
      v.appendChild(nt);
    }

    if (d.epilogue) {
      var ep = el('div', 'epilogue');
      var h = '<h3>' + esc(d.epilogue.title) + '</h3>';
      h += (d.epilogue.quotes || []).map(function (q) {
        return '<div class="quote">«' + esc(q.text) + '»' +
          (q.src ? '<span class="src">' + esc(q.src) + '</span>' : '') + '</div>';
      }).join('');
      if (d.epilogue.poem) h += '<div class="poem">' + d.epilogue.poem.map(esc).join('<br>') + '</div>';
      ep.innerHTML = h;
      v.appendChild(ep);
    }

    v.appendChild(renderActions(d, done));
  }

  function renderBlock(b, d, vals, locked) {
    var t = b.type;

    if (t === 'keywords')  return fieldKeywords(b, d, vals, locked);
    if (t === 'textarea')  return fieldTextarea(b, d, vals, locked);
    if (t === 'fields')    return fieldGroup(b, d, vals, locked);

    if (t === 'stepform') {
      var box = el('div', 'stepbox');
      box.innerHTML = '<span class="step-tag">' + esc(b.num) + '</span><h3>' + esc(b.title) + '</h3>' +
        (b.hint ? '<div class="body"><p>' + esc(b.hint) + '</p></div>' : '');
      var inner = renderBlock(b.field, d, vals, locked);
      if (inner) { inner.style.marginBottom = '0'; inner.style.marginTop = '14px'; box.appendChild(inner); }
      return box;
    }

    if (t === 'rule' || t === 'step') {
      var r = el('div', t === 'step' ? 'stepbox' : 'rule');
      r.innerHTML = (t === 'step' ? '<span class="step-tag">' + esc(b.num) + '</span>' : '') +
        '<h3>' + esc(b.title) + '</h3>' +
        '<div class="body">' + (b.body || []).map(function (p) { return '<p>' + p + '</p>'; }).join('') + '</div>';
      return r;
    }

    if (t === 'answerkey') {
      var a = el('div', 'answerkey');
      a.innerHTML = '<h3>' + esc(b.title) + '</h3><ul>' +
        b.items.map(function (i) { return '<li>' + i + '</li>'; }).join('') + '</ul>' +
        (b.after ? '<div class="after">' + b.after + '</div>' : '');
      return a;
    }

    if (t === 'callout') {
      var c = el('div', 'callout');
      c.innerHTML = '<h3>' + esc(b.title) + '</h3><ul>' +
        b.items.map(function (i) { return '<li>' + esc(i) + '</li>'; }).join('') + '</ul>';
      return c;
    }

    if (t === 'checklist') {
      var ch = el('div', 'checklist');
      ch.innerHTML = '<h3>' + esc(b.title) + '</h3><ul>' +
        b.items.map(function (i) { return '<li>' + esc(i) + '</li>'; }).join('') + '</ul>';
      return ch;
    }

    if (t === 'numbered') {
      var nb = el('div', 'numbered');
      nb.innerHTML = '<h3>' + esc(b.title) + '</h3><ol>' +
        b.items.map(function (i) { return '<li>' + esc(i) + '</li>'; }).join('') + '</ol>' +
        (b.after ? '<div class="after">' + esc(b.after) + '</div>' : '');
      return nb;
    }

    if (t === 'criteria') {
      var cr = el('div', 'criteria');
      cr.innerHTML = '<h3>' + esc(b.title) + '</h3>' + b.items.map(function (i) {
        return '<div class="criteria-item"><b>' + esc(i.name) + '</b><span>' + esc(i.text) + '</span></div>';
      }).join('');
      return cr;
    }

    if (t === 'truths') {
      var tr = el('div', 'truths');
      tr.innerHTML = '<h3>' + esc(b.title) + '</h3><div class="truths-grid">' +
        b.items.map(function (i) {
          return '<div class="truth"><b>' + esc(i.name) + '</b><ul>' +
            i.points.map(function (p) { return '<li>' + esc(p) + '</li>'; }).join('') + '</ul></div>';
        }).join('') + '</div>';
      return tr;
    }

    if (t === 'table3') {
      var tb = el('div', 'table3');
      tb.innerHTML = b.columns.map(function (c) {
        return '<div class="table3-col"><h4>' + esc(c.head) + '</h4><ul>' +
          c.rows.map(function (r) { return '<li>' + esc(r) + '</li>'; }).join('') +
          '</ul><div class="ref">' + esc(c.ref) + '</div></div>';
      }).join('');
      return tb;
    }

    return null;
  }

  function bindInput(node, dayId, key, locked) {
    if (locked) { node.disabled = true; return; }
    node.addEventListener('input', function () {
      setVal(dayId, key, node.value);
      if (node.tagName === 'TEXTAREA') autoGrow(node);
    });
  }
  function autoGrow(ta) {
    ta.style.height = 'auto';
    ta.style.height = Math.max(ta.scrollHeight, 44) + 'px';
  }

  function fieldKeywords(b, d, vals, locked) {
    var box = el('div', 'field');
    if (b.label) box.appendChild(el('div', 'field-label', esc(b.label)));
    for (var i = 0; i < b.count; i++) {
      (function (i) {
        var p = el('div', 'pair');

        var r1 = el('div', 'row');
        r1.appendChild(el('label', null, esc(b.termLabel) + ' ' + (i + 1)));
        var inp = el('input', 'input');
        inp.type = 'text';
        inp.placeholder = 'Слово или фраза из стиха';
        inp.value = vals[b.id + '.' + i + '.term'] || '';
        bindInput(inp, d.id, b.id + '.' + i + '.term', locked);
        r1.appendChild(inp);

        var r2 = el('div', 'row');
        r2.appendChild(el('label', null, 'Определение'));
        var ta = el('textarea', 'textarea');
        ta.rows = 2;
        ta.placeholder = 'Своими словами или из словаря';
        ta.value = vals[b.id + '.' + i + '.def'] || '';
        bindInput(ta, d.id, b.id + '.' + i + '.def', locked);
        r2.appendChild(ta);

        p.appendChild(r1); p.appendChild(r2);
        box.appendChild(p);
      })(i);
    }
    return box;
  }

  function fieldTextarea(b, d, vals, locked) {
    var box = el('div', 'field');
    if (b.label) box.appendChild(el('div', 'field-label', esc(b.label) +
      (b.optional ? '<span class="hint">Необязательно</span>' : '')));
    var ta = el('textarea', 'textarea');
    ta.rows = b.rows || 3;
    ta.placeholder = 'Ваш ответ…';
    ta.value = vals[b.id] || '';
    bindInput(ta, d.id, b.id, locked);
    box.appendChild(ta);
    return box;
  }

  function fieldGroup(b, d, vals, locked) {
    var box = el('div', 'field');
    if (b.label) box.appendChild(el('div', 'field-label', esc(b.label)));
    b.items.forEach(function (it) {
      var sf = el('div', 'subfield');
      sf.appendChild(el('label', null, esc(it.label)));
      var ta = el('textarea', 'textarea');
      ta.rows = b.big ? 4 : 2;
      ta.placeholder = 'Ваш ответ…';
      ta.value = vals[b.id + '.' + it.key] || '';
      bindInput(ta, d.id, b.id + '.' + it.key, locked);
      sf.appendChild(ta);
      box.appendChild(sf);
    });
    return box;
  }

  /* ───────────────────────── Кнопки внизу ───────────────────────── */

  function renderActions(d, done) {
    var wrap = el('div', 'actions');
    var inner = el('div', 'actions-inner');

    var i = idx(d.id);
    var next = C.order[i + 1];

    if (done) {
      if (next) {
        inner.appendChild(el('div', 'status', 'Ответы отправлены'));
        var nb = el('button', 'btn');
        nb.innerHTML = 'Следующий день ' + ICON.arrow;
        nb.onclick = function () { goDay(next); };
        inner.appendChild(nb);
      } else {
        inner.appendChild(el('div', 'status', 'Это последний день курса'));
      }
    } else {
      var st = el('div', 'status'); st.id = 'saveStatus';
      inner.appendChild(st);

      var btn = el('button', 'btn btn-lg');
      btn.id = 'submitBtn';
      btn.innerHTML = (next ? 'Отправить и продолжить ' : 'Завершить курс ') + ICON.arrow;
      btn.onclick = function () { submitDay(d, btn); };
      inner.appendChild(btn);
    }

    wrap.appendChild(inner);
    return wrap;
  }

  /* ───────────────────────── Отправка ───────────────────────── */

  function collectReadable(d, vals) {
    var out = [];
    function push(q, a) {
      a = (a || '').trim();
      if (a) out.push(q + '\n' + a);
    }
    function walk(blocks) {
      blocks.forEach(function (b) {
        if (b.type === 'stepform') { walkOne(b.field, b.num + '. ' + b.title); return; }
        walkOne(b, null);
      });
    }
    function walkOne(b, prefix) {
      var head = prefix ? prefix : (b.label || '');
      if (b.type === 'keywords') {
        for (var i = 0; i < b.count; i++) {
          var term = vals[b.id + '.' + i + '.term'], def = vals[b.id + '.' + i + '.def'];
          if ((term || '').trim() || (def || '').trim()) {
            push((head ? head + ' — ' : '') + b.termLabel + ' ' + (i + 1),
              (term || '') + ((def || '').trim() ? ' — ' + def : ''));
          }
        }
      } else if (b.type === 'textarea') {
        push(head || 'Ответ', vals[b.id]);
      } else if (b.type === 'fields') {
        b.items.forEach(function (it) {
          push((head ? head + ' · ' : '') + it.label, vals[b.id + '.' + it.key]);
        });
      }
    }
    walk(d.blocks || []);
    return out.join('\n\n');
  }

  function countFilled(d, vals) {
    var total = 0, filled = 0;
    function one(b) {
      if (b.type === 'keywords') {
        for (var i = 0; i < b.count; i++) {
          total += 1;
          if ((vals[b.id + '.' + i + '.term'] || '').trim() || (vals[b.id + '.' + i + '.def'] || '').trim()) filled++;
        }
      } else if (b.type === 'textarea') {
        if (b.optional) return;
        total += 1;
        if ((vals[b.id] || '').trim()) filled++;
      } else if (b.type === 'fields') {
        b.items.forEach(function (it) {
          total += 1;
          if ((vals[b.id + '.' + it.key] || '').trim()) filled++;
        });
      }
    }
    (d.blocks || []).forEach(function (b) { b.type === 'stepform' ? one(b.field) : one(b); });
    return { total: total, filled: filled };
  }

  function submitDay(d, btn) {
    var vals = valuesFor(d.id);
    var c = countFilled(d, vals);

    if (c.total > 0 && c.filled === 0) {
      toast('Заполните хотя бы одно поле');
      var f = $('#view .input, #view .textarea');
      if (f) f.focus();
      return;
    }
    if (c.filled < c.total) {
      if (!confirm('Заполнено ' + c.filled + ' из ' + c.total + ' полей.\n\nОтправить ответы и перейти к следующему дню?\nПосле отправки изменить ответы будет нельзя.')) return;
    }

    clearTimeout(S.syncTimer);
    btn.disabled = true;
    btn.innerHTML = 'Отправляем…';
    flagStatus('saving');

    API.post({
      action: 'submit',
      token: S.token,
      dayId: d.id,
      n: d.n,
      module: d.moduleNum + '. ' + d.moduleTitle,
      title: 'День ' + d.dayInModule + '. ' + d.title,
      total: C.totalDays,
      answers: vals,
      readable: collectReadable(d, vals)
    }).then(function (res) {
      if (!res || !res.ok) throw new Error((res && res.error) || 'fail');
      S.submitted[d.id] = { answers: vals, submittedAt: res.submittedAt };
      var next = C.order[idx(d.id) + 1];
      paintNav(); paintTop();
      if (next) { goDay(next); toast('День отправлен наставнику'); }
      else { renderDay(); toast('Поздравляем — курс пройден!'); }
    }).catch(function () {
      btn.disabled = false;
      btn.innerHTML = 'Отправить ещё раз ' + ICON.arrow;
      flagStatus('error');
      var st = $('#saveStatus');
      if (st) st.textContent = 'Нет связи. Ответы сохранены — попробуйте ещё раз.';
    });
  }

  /* ───────────────────────── Приложение ───────────────────────── */

  function goAppendix() {
    S.view = 'appendix';
    renderAppendix(); paintNav(); paintTop();
    window.scrollTo(0, 0);
  }

  function renderAppendix() {
    var v = $('#view');
    var h = '<div class="day-head"><div class="eyebrow">Приложение</div>' +
            '<h1>Материалы для благовестия</h1>' +
            '<div class="sub">Справочный раздел — заполнять ничего не нужно</div></div><div class="appx">';
    C.appendix.forEach(function (a) {
      h += '<h2>' + esc(a.title) + '</h2>';
      a.sections.forEach(function (s) {
        if (s.title && s.type !== 'compare') h += '<h3>' + esc(s.title) + '</h3>';
        if (s.type === 'list') h += '<ul>' + s.items.map(function (i) { return '<li>' + i + '</li>'; }).join('') + '</ul>';
        else if (s.type === 'numbered') h += '<ol>' + s.items.map(function (i) { return '<li>' + esc(i) + '</li>'; }).join('') + '</ol>' +
          (s.after ? '<p style="color:var(--muted);font-size:14.5px">' + esc(s.after) + '</p>' : '');
        else if (s.type === 'text') h += s.items.map(function (i) { return '<p>' + esc(i) + '</p>'; }).join('') +
          (s.src ? '<div class="src">' + esc(s.src) + '</div>' : '');
        else if (s.type === 'qa') h += '<div class="qa"><b>' + esc(s.q) + '</b><ul>' +
          s.a.map(function (i) { return '<li>' + esc(i) + '</li>'; }).join('') + '</ul></div>';
        else if (s.type === 'compare') {
          h += '<h3>' + esc(s.title) + '</h3><div class="cmp">' + s.columns.map(function (c) {
            return '<div><h4>' + esc(c.head) + '</h4><ul>' +
              c.items.map(function (i) { return '<li>' + i + '</li>'; }).join('') + '</ul></div>';
          }).join('') + '</div>';
        }
      });
    });
    h += '</div>';
    v.innerHTML = h;
  }

  /* ───────────────────────── Запуск ───────────────────────── */

  function start() {
    initTheme();

    var p = new URLSearchParams(location.search);
    var token = (p.get('s') || '').trim();
    if (!token) { try { token = localStorage.getItem('krsh:last') || ''; } catch (e) {} }
    if (!token) { screenWelcome(); return; }
    S.token = token;
    try { localStorage.setItem('krsh:last', token); } catch (e) {}

    if (!API.isConfigured()) {
      screenError('Сайт ещё не настроен',
        'В файле <code>assets/config.js</code> нужно указать адрес веб-приложения Google Apps Script. Инструкция — в README.md.');
      return;
    }

    screenLoading();

    API.get({ action: 'student', token: token }).then(function (res) {
      if (!res || !res.ok) {
        if (res && res.error === 'not_found') {
          screenWelcome('Ссылка не найдена. Проверьте её у наставника.');
        } else {
          screenError('Не удалось загрузить курс', 'Попробуйте обновить страницу через минуту.', true);
        }
        return;
      }
      S.student = res.student;
      S.submitted = res.submitted || {};

      // черновики: сервер + локальные (локальные приоритетнее — они свежее)
      S.values = {};
      var remote = res.drafts || {};
      Object.keys(remote).forEach(function (k) { S.values[k] = remote[k].answers || {}; });
      var local = loadLocal();
      Object.keys(local).forEach(function (k) {
        S.values[k] = Object.assign({}, S.values[k] || {}, local[k] || {});
      });
      saveLocal();

      shell();
      var hash = (location.hash || '').replace('#', '');
      S.currentId = (DAYS[hash] && isUnlocked(hash)) ? hash : firstOpenDay();
      goDay(S.currentId);
    }).catch(function () {
      screenError('Нет связи с сервером',
        'Проверьте интернет и попробуйте ещё раз. Уже введённые ответы сохранены на устройстве.', true);
    });
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start);
  else start();
})();
