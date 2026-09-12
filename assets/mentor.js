/* ============================================================
   Курс подготовки ко крещению — панель наставника
   ============================================================ */
(function () {
  'use strict';

  var C = window.COURSE;
  var DAYS = {};
  C.modules.forEach(function (m) { m.days.forEach(function (d) { DAYS[d.id] = d; }); });

  var S = { key: '', mentor: null, students: [], filter: '', open: null, openDay: null, cache: {} };

  var $ = function (s, r) { return (r || document).querySelector(s); };
  function el(t, c, h) { var n = document.createElement(t); if (c) n.className = c; if (h != null) n.innerHTML = h; return n; }
  function esc(s) {
    return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;')
      .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }
  function initials(name) {
    var p = String(name).trim().split(/\s+/);
    return ((p[0] || '')[0] || '?').toUpperCase() + ((p[1] || '')[0] || '').toUpperCase();
  }
  function fmt(iso) {
    if (!iso) return '';
    var d = new Date(iso);
    if (isNaN(d)) return '';
    var today = new Date(); today.setHours(0, 0, 0, 0);
    var dd = new Date(d); dd.setHours(0, 0, 0, 0);
    var diff = Math.round((today - dd) / 86400000);
    if (diff === 0) return 'сегодня, ' + d.toLocaleTimeString('ru-RU', { hour: '2-digit', minute: '2-digit' });
    if (diff === 1) return 'вчера, ' + d.toLocaleTimeString('ru-RU', { hour: '2-digit', minute: '2-digit' });
    if (diff < 7) return diff + ' дн. назад';
    return d.toLocaleDateString('ru-RU', { day: 'numeric', month: 'short', year: 'numeric' });
  }
  var tt;
  function toast(m) {
    var t = $('#toast'); t.textContent = m; t.classList.add('show');
    clearTimeout(tt); tt = setTimeout(function () { t.classList.remove('show'); }, 2400);
  }

  function initTheme() {
    var saved = null; try { saved = localStorage.getItem('krsh:theme'); } catch (e) {}
    if (saved) document.documentElement.setAttribute('data-theme', saved);
  }
  function paintThemeBtn() {
    var b = $('#themeBtn'); if (!b) return;
    var dark = document.documentElement.getAttribute('data-theme') === 'dark' ||
      (!document.documentElement.getAttribute('data-theme') && window.matchMedia('(prefers-color-scheme: dark)').matches);
    b.innerHTML = dark ? ICON.sun : ICON.moon;
  }
  function toggleTheme() {
    var cur = document.documentElement.getAttribute('data-theme');
    var dark = cur === 'dark' || (!cur && window.matchMedia('(prefers-color-scheme: dark)').matches);
    var n = dark ? 'light' : 'dark';
    document.documentElement.setAttribute('data-theme', n);
    try { localStorage.setItem('krsh:theme', n); } catch (e) {}
    paintThemeBtn();
  }

  function centered(html) {
    document.body.innerHTML = '<div class="centered"><div class="card">' + html + '</div></div>' +
      '<div class="toast" id="toast"></div>';
  }

  function screenKey(err) {
    centered(
      '<div class="mark-lg">' + ICON.users + '</div>' +
      '<h1>Панель наставника</h1>' +
      '<p>Введите свой ключ доступа. Его можно взять на листе «Наставники» в таблице курса.</p>' +
      (err ? '<p style="color:var(--danger);margin-top:14px">' + esc(err) + '</p>' : '') +
      '<div class="form"><input class="input" id="k" placeholder="Ключ наставника" autocomplete="off" spellcheck="false">' +
      '<button class="btn btn-lg btn-block" id="go" style="margin-top:12px">Войти</button></div>'
    );
    var go = function () {
      var v = $('#k').value.trim();
      var m = v.match(/[?&]k=([^&\s]+)/);
      if (m) v = decodeURIComponent(m[1]);
      if (v) location.search = '?k=' + encodeURIComponent(v);
    };
    $('#go').onclick = go;
    $('#k').onkeydown = function (e) { if (e.key === 'Enter') go(); };
  }

  /* ───────────────────────── Каркас ───────────────────────── */

  function shell() {
    document.body.innerHTML =
      '<header class="topbar">' +
      '  <div class="topbar-title"><strong>Панель наставника</strong><span id="tbSub"></span></div>' +
      '  <button class="icon-btn" id="reloadBtn" aria-label="Обновить">' + ICON.refresh + '</button>' +
      '  <button class="icon-btn" id="themeBtn"></button>' +
      '</header>' +
      '<main class="main"><div class="wrap wrap-wide">' +
      '  <div class="m-head"><h1 id="hello"></h1><div class="sub" id="sub2"></div></div>' +
      '  <div class="stats" id="stats"></div>' +
      '  <div class="toolbar">' +
      '    <div class="search">' + ICON.search +
      '      <input class="input" id="q" placeholder="Поиск по имени" autocomplete="off"></div>' +
      '  </div>' +
      '  <div class="s-list" id="list"></div>' +
      '</div></main>' +
      '<div class="sheet-back" id="sheetBack"></div>' +
      '<aside class="sheet" id="sheet">' +
      '  <div class="sheet-head">' +
      '    <button class="icon-btn" id="sheetClose" aria-label="Закрыть">' + ICON.close + '</button>' +
      '    <div class="t"><strong id="shName"></strong><span id="shMeta"></span></div>' +
      '  </div>' +
      '  <div class="sheet-body" id="sheetBody"></div>' +
      '</aside>' +
      '<div class="toast" id="toast"></div>';

    $('#themeBtn').onclick = toggleTheme;
    $('#reloadBtn').onclick = function () { load(true); };
    $('#sheetClose').onclick = closeSheet;
    $('#sheetBack').onclick = closeSheet;
    $('#q').oninput = function () { S.filter = this.value.toLowerCase().trim(); paintList(); };
    document.addEventListener('keydown', function (e) { if (e.key === 'Escape') closeSheet(); });
    paintThemeBtn();
  }

  function paintHead() {
    var total = S.students.length;
    var finished = S.students.filter(function (s) { return Object.keys(s.days).length >= C.totalDays; }).length;
    var active = S.students.filter(function (s) {
      if (!s.lastActivity) return false;
      return (Date.now() - new Date(s.lastActivity).getTime()) < 7 * 86400000;
    }).length;
    var sum = S.students.reduce(function (a, s) { return a + Object.keys(s.days).length; }, 0);

    $('#hello').textContent = S.mentor && S.mentor.name ? S.mentor.name : 'Ученики';
    $('#sub2').textContent = total === 0 ? 'Пока нет учеников' :
      'Учеников: ' + total + ' · дней пройдено всего: ' + sum;
    $('#tbSub').textContent = total + ' ' + plural(total, 'ученик', 'ученика', 'учеников');

    $('#stats').innerHTML =
      stat(total, 'Всего учеников') +
      stat(active, 'Активны за неделю') +
      stat(finished, 'Завершили курс') +
      stat(total ? Math.round(sum / total) : 0, 'Дней в среднем');
  }
  function stat(v, k) { return '<div class="stat"><div class="v">' + v + '</div><div class="k">' + k + '</div></div>'; }
  function plural(n, a, b, c) {
    var n10 = n % 10, n100 = n % 100;
    if (n10 === 1 && n100 !== 11) return a;
    if (n10 >= 2 && n10 <= 4 && (n100 < 10 || n100 >= 20)) return b;
    return c;
  }

  function paintList() {
    var list = $('#list');
    list.innerHTML = '';
    var items = S.students.filter(function (s) {
      return !S.filter || s.name.toLowerCase().indexOf(S.filter) >= 0;
    });
    if (!items.length) {
      list.appendChild(el('div', 'empty', S.students.length
        ? 'Никого не нашли по запросу «' + esc(S.filter) + '»'
        : 'Добавьте первого ученика в таблице: меню «📖 Курс → ➕ Добавить ученика»'));
      return;
    }
    items.sort(function (a, b) {
      return (b.lastActivity || '').localeCompare(a.lastActivity || '') || a.name.localeCompare(b.name);
    });
    items.forEach(function (s) {
      var n = Object.keys(s.days).length;
      var card = el('div', 's-card');
      var dots = C.order.map(function (id, i) {
        var mod = DAYS[id].dayInModule === DAYS[id].daysInModule ? ' sep' : '';
        return '<span class="dot-day' + (s.days[id] ? ' done' : '') + mod + '" title="' +
               esc(DAYS[id].moduleTitle + ' · ' + DAYS[id].title) + '"></span>';
      }).join('');
      card.innerHTML =
        '<div class="s-top">' +
        '  <div class="avatar">' + esc(initials(s.name)) + '</div>' +
        '  <div class="s-name"><strong>' + esc(s.name) + '</strong>' +
        '    <span>' + (s.lastActivity ? esc(fmt(s.lastActivity)) : 'ещё не начинал') +
             (s.mentor ? ' · ' + esc(s.mentor) : '') + '</span></div>' +
        '  <div class="s-count' + (n >= C.totalDays ? ' full' : '') + '">' + n + '/' + C.totalDays + '</div>' +
        '</div>' +
        '<div class="progress-bar"><i style="width:' + Math.round(n / C.totalDays * 100) + '%"></i></div>' +
        '<div class="dots">' + dots + '</div>';
      card.onclick = function () { openStudent(s); };
      list.appendChild(card);
    });
  }

  /* ───────────────────────── Карточка ученика ───────────────────────── */

  function openStudent(s) {
    S.open = s;
    S.openDay = null;
    $('#shName').textContent = s.name;
    $('#shMeta').textContent = Object.keys(s.days).length + ' из ' + C.totalDays + ' дней' +
      (s.lastActivity ? ' · ' + fmt(s.lastActivity) : '');
    $('#sheet').classList.add('open');
    $('#sheetBack').classList.add('show');
    document.body.style.overflow = 'hidden';
    paintSheet(null);

    if (S.cache[s.token]) { paintSheet(null); return; }
    API.get({ action: 'mentorStudent', key: S.key, token: s.token }).then(function (res) {
      if (res && res.ok) { S.cache[s.token] = res.submitted || {}; if (S.open === s) paintSheet(S.openDay); }
    }).catch(function () {});
  }

  function closeSheet() {
    $('#sheet').classList.remove('open');
    $('#sheetBack').classList.remove('show');
    document.body.style.overflow = '';
    S.open = null;
  }

  function paintSheet(activeDay) {
    var s = S.open;
    if (!s) return;
    var body = $('#sheetBody');
    body.innerHTML = '';

    // Ссылка
    var linkRow = el('div', 'toolbar');
    var copy = el('button', 'btn btn-ghost');
    copy.innerHTML = ICON.copy + ' Скопировать ссылку';
    copy.onclick = function () {
      var link = s.link || (location.origin + location.pathname.replace(/mentor\.html$/, '') + '?s=' + s.token);
      navigator.clipboard.writeText(link).then(
        function () { toast('Ссылка скопирована'); },
        function () { prompt('Скопируйте ссылку:', link); }
      );
    };
    linkRow.appendChild(copy);
    var openBtn = el('a', 'btn btn-ghost');
    openBtn.href = s.link || ('?s=' + s.token);
    openBtn.target = '_blank'; openBtn.rel = 'noopener';
    openBtn.innerHTML = 'Открыть как ученик';
    linkRow.appendChild(openBtn);

    var allBtn = el('button', 'btn btn-ghost');
    allBtn.innerHTML = (activeDay === '*' ? 'Свернуть' : 'Все ответы подряд');
    allBtn.onclick = function () { S.openDay = (S.openDay === '*' ? null : '*'); paintSheet(S.openDay); };
    linkRow.appendChild(allBtn);

    if (activeDay) {
      var prBtn = el('button', 'btn btn-ghost');
      prBtn.innerHTML = ICON.print + ' Печать';
      prBtn.onclick = function () { window.print(); };
      linkRow.appendChild(prBtn);
    }
    body.appendChild(linkRow);

    // Дни по модулям
    var mods = el('div', 'd-mods');
    C.modules.forEach(function (m) {
      var row = el('div', 'd-mod');
      var doneN = m.days.filter(function (d) { return !!s.days[d.id]; }).length;
      row.appendChild(el('div', 'nm',
        '<b>' + m.num + '. ' + esc(m.title) + '</b><i>' + doneN + ' из ' + m.days.length + '</i>'));
      var grid = el('div', 'd-grid');
      m.days.forEach(function (d) {
        var done = !!s.days[d.id];
        var b = el('button', 'd-cell' + (done ? ' done' : '') + (activeDay === d.id ? ' active' : ''));
        b.textContent = d.dayInModule;
        b.title = d.title + (done ? '' : ' — не пройден');
        b.disabled = !done;
        b.onclick = function () { S.openDay = (S.openDay === d.id ? null : d.id); paintSheet(S.openDay); };
        grid.appendChild(b);
      });
      row.appendChild(grid);
      mods.appendChild(row);
    });
    body.appendChild(mods);

    if (!activeDay) {
      var n = Object.keys(s.days).length;
      body.appendChild(el('div', 'empty', n
        ? 'Нажмите на любой пройденный день, чтобы прочитать ответы'
        : 'Ученик ещё не отправил ни одного дня'));
      return;
    }

    var answers = S.cache[s.token];
    if (!answers) { body.appendChild(el('div', 'empty', '<div class="spinner"></div>')); return; }

    if (activeDay === '*') {
      var ids = C.order.filter(function (id) { return !!answers[id]; });
      if (!ids.length) { body.appendChild(el('div', 'empty', 'Пока нет отправленных дней')); return; }
      ids.forEach(function (id, k) {
        var dd = DAYS[id], rr = answers[id];
        var blk = el('div', 'ans-block');
        var hh = el('div', 'ans-head');
        hh.innerHTML = '<h3>' + esc(dd.title) + '</h3><div class="meta">' +
          esc(dd.moduleNum + '. ' + dd.moduleTitle) + ' · день ' + dd.dayInModule +
          ' · отправлено ' + esc(fmt(rr.submittedAt)) + '</div>';
        blk.appendChild(hh);
        var w = el('div');
        renderAnswers(dd, rr.answers || {}, w);
        blk.appendChild(w);
        body.appendChild(blk);
      });
      return;
    }

    var rec = answers[activeDay];
    var d = DAYS[activeDay];
    var head = el('div', 'ans-head');
    head.innerHTML = '<h3>' + esc(d.title) + '</h3><div class="meta">' +
      esc(d.moduleNum + '. ' + d.moduleTitle) + ' · день ' + d.dayInModule +
      (rec ? ' · отправлено ' + esc(fmt(rec.submittedAt)) : '') + '</div>';
    body.appendChild(head);

    if (!rec) { body.appendChild(el('div', 'empty', 'Ответы не найдены')); return; }

    var vals = rec.answers || {};
    var wrap = el('div');
    renderAnswers(d, vals, wrap);
    body.appendChild(wrap);
  }

  function renderAnswers(d, vals, wrap) {
    var any = false;
    function item(q, a, def) {
      var it = el('div', 'ans-item');
      var txt = (a || '').trim();
      var dtx = (def || '').trim();
      if (txt || dtx) any = true;
      var html;
      if (!txt && !dtx) html = '<div class="ans-a empty-a">— не заполнено</div>';
      else if (def !== undefined) html = '<div class="ans-a"><span class="term">' + esc(txt || '—') + '</span>' +
        (dtx ? '<span class="def">' + esc(dtx) + '</span>' : '') + '</div>';
      else html = '<div class="ans-a">' + esc(txt) + '</div>';
      it.innerHTML = '<div class="ans-q">' + esc(q) + '</div>' + html;
      wrap.appendChild(it);
    }
    function one(b, prefix) {
      var head = prefix || b.label || '';
      if (b.type === 'keywords') {
        for (var i = 0; i < b.count; i++) {
          var t = (vals[b.id + '.' + i + '.term'] || '').trim();
          var def = (vals[b.id + '.' + i + '.def'] || '').trim();
          item((head ? head + ' — ' : '') + b.termLabel + ' ' + (i + 1), t, def);
        }
      } else if (b.type === 'textarea') {
        item(head || 'Ответ', vals[b.id]);
      } else if (b.type === 'fields') {
        b.items.forEach(function (it2) {
          item((head ? head + ' · ' : '') + it2.label, vals[b.id + '.' + it2.key]);
        });
      }
    }
    (d.blocks || []).forEach(function (b) {
      if (b.type === 'stepform') one(b.field, b.num + '. ' + b.title);
      else one(b, null);
    });
    if (!wrap.children.length) wrap.appendChild(el('div', 'empty', 'В этом дне нет полей для заполнения'));
  }

  /* ───────────────────────── Загрузка ───────────────────────── */

  function load(manual) {
    API.get({ action: 'mentor', key: S.key }).then(function (res) {
      if (!res || !res.ok) {
        if (res && res.error === 'not_found') screenKey('Ключ не найден. Проверьте лист «Наставники».');
        else centered('<h1>Ошибка</h1><p>Не удалось загрузить данные.</p>');
        return;
      }
      S.mentor = res.mentor;
      S.students = res.students || [];
      S.cache = {};
      if (!$('#list')) shell();
      paintHead(); paintList();
      if (manual) toast('Обновлено');
    }).catch(function () {
      centered('<h1>Нет связи</h1><p>Проверьте интернет и обновите страницу.</p>');
    });
  }

  function start() {
    initTheme();
    var p = new URLSearchParams(location.search);
    var key = (p.get('k') || '').trim();
    if (!key) { try { key = localStorage.getItem('krsh:mentorKey') || ''; } catch (e) {} }
    if (!key) { screenKey(); return; }
    S.key = key;
    try { localStorage.setItem('krsh:mentorKey', key); } catch (e) {}

    if (!API.isConfigured()) {
      centered('<h1>Сайт ещё не настроен</h1><p>В файле <code>assets/config.js</code> нужно указать адрес веб-приложения Google Apps Script.</p>');
      return;
    }
    centered('<div class="spinner"></div><p style="margin-top:18px">Загружаем данные…</p>');
    load(false);
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start);
  else start();
})();
