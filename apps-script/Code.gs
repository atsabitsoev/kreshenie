/**
 * Курс подготовки ко крещению — бэкенд на Google Apps Script.
 *
 * Хранилище: эта же Google-таблица.
 * Фронтенд: статический сайт на GitHub Pages.
 *
 * Установка описана в README.md.
 */

/* ───────────────────────────── Константы ───────────────────────────── */

var SH_STUDENTS = 'Ученики';
var SH_ANSWERS  = 'Ответы';
var SH_DRAFTS   = 'Черновики';
var SH_MENTORS  = 'Наставники';
var SH_CONFIG   = 'Настройки';

var STUDENT_HEADERS = [
  'Имя', 'Контакт', 'Наставник', 'Токен', 'Ссылка для ученика',
  'Создан', 'Прогресс', 'Последняя активность', 'Заметки наставника'
];
var ANSWER_HEADERS = [
  'Токен', 'Ученик', 'День', '№', 'Модуль', 'Тема', 'Отправлено', 'Ответы', 'JSON'
];
var DRAFT_HEADERS = ['Токен', 'День', 'Изменён', 'JSON'];
var MENTOR_HEADERS = ['Имя', 'Ключ', 'Ссылка на панель'];

/* ──────────────────────────── Меню таблицы ─────────────────────────── */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📖 Курс')
    .addItem('① Настроить таблицу', 'setupSheets')
    .addItem('② Указать адрес сайта', 'promptSiteUrl')
    .addSeparator()
    .addItem('➕ Добавить ученика', 'promptAddStudent')
    .addItem('🔑 Добавить наставника', 'promptAddMentor')
    .addSeparator()
    .addItem('🔗 Выдать ссылки всем без токена', 'fillMissingTokens')
    .addItem('📋 Скопировать ссылку выбранного ученика', 'showSelectedLink')
    .addItem('🔄 Пересчитать прогресс', 'recalcProgress')
    .addToUi();
}

/* ─────────────────────────── Инициализация ─────────────────────────── */

function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var st = ensureSheet_(ss, SH_STUDENTS, STUDENT_HEADERS);
  st.setColumnWidth(1, 200); st.setColumnWidth(2, 160); st.setColumnWidth(3, 150);
  st.setColumnWidth(4, 190); st.setColumnWidth(5, 360); st.setColumnWidth(6, 140);
  st.setColumnWidth(7, 110); st.setColumnWidth(8, 150); st.setColumnWidth(9, 260);

  var an = ensureSheet_(ss, SH_ANSWERS, ANSWER_HEADERS);
  an.setColumnWidth(1, 150); an.setColumnWidth(2, 170); an.setColumnWidth(3, 60);
  an.setColumnWidth(4, 50);  an.setColumnWidth(5, 130); an.setColumnWidth(6, 230);
  an.setColumnWidth(7, 150); an.setColumnWidth(8, 640); an.setColumnWidth(9, 120);
  an.hideColumns(9);

  var dr = ensureSheet_(ss, SH_DRAFTS, DRAFT_HEADERS);
  dr.hideSheet();

  var mn = ensureSheet_(ss, SH_MENTORS, MENTOR_HEADERS);
  mn.setColumnWidth(1, 200); mn.setColumnWidth(2, 190); mn.setColumnWidth(3, 380);

  var cf = ensureSheet_(ss, SH_CONFIG, ['Параметр', 'Значение']);
  if (cf.getLastRow() < 2) {
    cf.getRange(2, 1, 1, 2).setValues([['SITE_URL', '']]);
    cf.getRange(3, 1, 1, 2).setValues([['Подсказка', 'Укажите в SITE_URL адрес сайта, например https://user.github.io/kreshenie/']]);
  }
  cf.setColumnWidth(1, 160); cf.setColumnWidth(2, 520);

  SpreadsheetApp.getUi().alert('Готово', 'Листы созданы и настроены.\n\nДальше: «📖 Курс → ② Указать адрес сайта».', SpreadsheetApp.getUi().ButtonSet.OK);
}

function ensureSheet_(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  var existing = sh.getRange(1, 1, 1, Math.max(headers.length, sh.getLastColumn() || 1)).getValues()[0];
  var needsHeader = false;
  for (var i = 0; i < headers.length; i++) {
    if (String(existing[i] || '').trim() !== headers[i]) { needsHeader = true; break; }
  }
  if (needsHeader) sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold').setBackground('#111827').setFontColor('#ffffff')
    .setVerticalAlignment('middle');
  sh.setFrozenRows(1);
  sh.setRowHeight(1, 34);
  return sh;
}

/* ─────────────────────────── Настройки сайта ───────────────────────── */

function getSiteUrl_() {
  var cf = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_CONFIG);
  if (!cf) return '';
  var vals = cf.getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === 'SITE_URL') return String(vals[i][1] || '').trim();
  }
  return '';
}

function promptSiteUrl() {
  var ui = SpreadsheetApp.getUi();
  var cur = getSiteUrl_();
  var res = ui.prompt('Адрес сайта',
    'Введите адрес опубликованного сайта (GitHub Pages).\nНапример: https://username.github.io/kreshenie/\n\nСейчас: ' + (cur || '— не задан —'),
    ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var url = res.getResponseText().trim();
  if (!url) return;
  if (url.slice(-1) !== '/') url += '/';

  var cf = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_CONFIG);
  var vals = cf.getDataRange().getValues();
  var row = -1;
  for (var i = 1; i < vals.length; i++) if (String(vals[i][0]).trim() === 'SITE_URL') row = i + 1;
  if (row === -1) { row = cf.getLastRow() + 1; cf.getRange(row, 1).setValue('SITE_URL'); }
  cf.getRange(row, 2).setValue(url);

  refreshAllLinks_();
  ui.alert('Готово', 'Адрес сохранён, все ссылки обновлены.', ui.ButtonSet.OK);
}

/* ──────────────────────── Ученики и наставники ─────────────────────── */

function makeToken_(prefix) {
  var alphabet = 'abcdefghijkmnpqrstuvwxyz23456789';
  var s = '';
  for (var i = 0; i < 18; i++) s += alphabet.charAt(Math.floor(Math.random() * alphabet.length));
  return prefix + s;
}

function studentLink_(token) {
  var base = getSiteUrl_();
  if (!base) return '(сначала укажите адрес сайта в меню «Курс»)';
  return base + '?s=' + token;
}
function mentorLink_(key) {
  var base = getSiteUrl_();
  if (!base) return '(сначала укажите адрес сайта в меню «Курс»)';
  return base + 'mentor.html?k=' + key;
}

function promptAddStudent() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('Новый ученик', 'Имя и фамилия ученика:', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var name = res.getResponseText().trim();
  if (!name) return;

  var res2 = ui.prompt('Наставник', 'Имя наставника (можно оставить пустым):', ui.ButtonSet.OK_CANCEL);
  var mentor = res2.getSelectedButton() === ui.Button.OK ? res2.getResponseText().trim() : '';

  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_STUDENTS);
  var token = makeToken_('s-');
  var link = studentLink_(token);
  sh.appendRow([name, '', mentor, token, link, new Date(), '0 / 31', '', '']);
  var row = sh.getLastRow();
  sh.getRange(row, 6).setNumberFormat('dd.MM.yyyy HH:mm');

  ui.alert('Ученик создан',
    name + '\n\nИндивидуальная ссылка:\n' + link + '\n\nСсылка также записана в таблицу (столбец «Ссылка для ученика»).',
    ui.ButtonSet.OK);
}

function promptAddMentor() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('Новый наставник', 'Имя наставника:', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var name = res.getResponseText().trim();
  if (!name) return;
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_MENTORS);
  var key = makeToken_('m-');
  var link = mentorLink_(key);
  sh.appendRow([name, key, link]);
  ui.alert('Наставник создан', name + '\n\nСсылка на панель наставника:\n' + link, ui.ButtonSet.OK);
}

/** Автоматически выдаёт токен и ссылку, как только в столбце «Имя» появилось значение. */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    var sh = e.range.getSheet();
    if (sh.getName() !== SH_STUDENTS) return;
    var row = e.range.getRow();
    if (row < 2) return;
    if (e.range.getColumn() !== 1) return;
    var name = String(sh.getRange(row, 1).getValue()).trim();
    if (!name) return;
    if (String(sh.getRange(row, 4).getValue()).trim()) return; // токен уже есть
    var token = makeToken_('s-');
    sh.getRange(row, 4).setValue(token);
    sh.getRange(row, 5).setValue(studentLink_(token));
    sh.getRange(row, 6).setValue(new Date()).setNumberFormat('dd.MM.yyyy HH:mm');
    sh.getRange(row, 7).setValue('0 / 31');
  } catch (err) { /* тихо */ }
}

function fillMissingTokens() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_STUDENTS);
  var last = sh.getLastRow();
  var added = 0;
  for (var r = 2; r <= last; r++) {
    var name = String(sh.getRange(r, 1).getValue()).trim();
    if (!name) continue;
    if (String(sh.getRange(r, 4).getValue()).trim()) continue;
    var token = makeToken_('s-');
    sh.getRange(r, 4).setValue(token);
    sh.getRange(r, 5).setValue(studentLink_(token));
    if (!sh.getRange(r, 6).getValue()) sh.getRange(r, 6).setValue(new Date()).setNumberFormat('dd.MM.yyyy HH:mm');
    added++;
  }
  refreshAllLinks_();
  SpreadsheetApp.getUi().alert('Готово', 'Выдано новых ссылок: ' + added, SpreadsheetApp.getUi().ButtonSet.OK);
}

function refreshAllLinks_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var st = ss.getSheetByName(SH_STUDENTS);
  if (st && st.getLastRow() > 1) {
    var n = st.getLastRow() - 1;
    var tokens = st.getRange(2, 4, n, 1).getValues();
    var links = tokens.map(function (t) {
      var v = String(t[0]).trim();
      return [v ? studentLink_(v) : ''];
    });
    st.getRange(2, 5, n, 1).setValues(links);
  }
  var mn = ss.getSheetByName(SH_MENTORS);
  if (mn && mn.getLastRow() > 1) {
    var m = mn.getLastRow() - 1;
    var keys = mn.getRange(2, 2, m, 1).getValues();
    var mlinks = keys.map(function (t) {
      var v = String(t[0]).trim();
      return [v ? mentorLink_(v) : ''];
    });
    mn.getRange(2, 3, m, 1).setValues(mlinks);
  }
}

function showSelectedLink() {
  var ui = SpreadsheetApp.getUi();
  var sh = SpreadsheetApp.getActiveSheet();
  if (sh.getName() !== SH_STUDENTS) { ui.alert('Откройте лист «Ученики» и выберите строку ученика.'); return; }
  var row = sh.getActiveRange().getRow();
  if (row < 2) { ui.alert('Выберите строку с учеником.'); return; }
  var name = sh.getRange(row, 1).getValue();
  var link = sh.getRange(row, 5).getValue();
  ui.alert('Ссылка ученика', name + '\n\n' + link, ui.ButtonSet.OK);
}

function recalcProgress() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var st = ss.getSheetByName(SH_STUDENTS);
  var an = ss.getSheetByName(SH_ANSWERS);
  if (!st || st.getLastRow() < 2) return;
  var counts = {}, lastAt = {};
  if (an && an.getLastRow() > 1) {
    var rows = an.getRange(2, 1, an.getLastRow() - 1, 7).getValues();
    for (var i = 0; i < rows.length; i++) {
      var tk = String(rows[i][0]).trim();
      if (!tk) continue;
      counts[tk] = (counts[tk] || 0) + 1;
      var d = rows[i][6];
      if (d instanceof Date && (!lastAt[tk] || d > lastAt[tk])) lastAt[tk] = d;
    }
  }
  var n = st.getLastRow() - 1;
  var tokens = st.getRange(2, 4, n, 1).getValues();
  var out = tokens.map(function (t) {
    var tk = String(t[0]).trim();
    var c = counts[tk] || 0;
    var la = lastAt[tk] ? Utilities.formatDate(lastAt[tk], Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm') : '';
    return [c + ' / 31', la];
  });
  st.getRange(2, 7, n, 2).setValues(out);
  SpreadsheetApp.getUi().alert('Готово', 'Прогресс пересчитан.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/* ──────────────────────────── Веб-приложение ───────────────────────── */

function doGet(e)  { return handle_(e, e && e.parameter ? e.parameter : {}); }
function doPost(e) {
  var body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) { body = {}; }
  var params = {};
  if (e && e.parameter) for (var k in e.parameter) params[k] = e.parameter[k];
  for (var k2 in body) params[k2] = body[k2];
  return handle_(e, params);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function handle_(e, p) {
  try {
    var action = String(p.action || '');
    if (action === 'ping')          return json_({ ok: true, version: 1 });
    if (action === 'student')       return json_(apiStudent_(p));
    if (action === 'submit')        return json_(apiSubmit_(p));
    if (action === 'draft')         return json_(apiDraft_(p));
    if (action === 'mentor')        return json_(apiMentor_(p));
    if (action === 'mentorStudent') return json_(apiMentorStudent_(p));
    return json_({ ok: false, error: 'unknown_action' });
  } catch (err) {
    return json_({ ok: false, error: 'server_error', detail: String(err) });
  }
}

/* ── Ученик ── */

function findStudent_(token) {
  if (!token) return null;
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_STUDENTS);
  if (!sh || sh.getLastRow() < 2) return null;
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][3]).trim() === token) {
      return { row: i + 2, name: String(rows[i][0]), contact: String(rows[i][1]),
               mentor: String(rows[i][2]), token: token, createdAt: rows[i][5] };
    }
  }
  return null;
}

function answersFor_(token) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_ANSWERS);
  var out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim() !== token) continue;
    var dayId = String(rows[i][2]).trim();
    var data = {};
    try { data = JSON.parse(rows[i][8] || '{}'); } catch (err) { data = {}; }
    out[dayId] = { answers: data, submittedAt: toIso_(rows[i][6]), row: i + 2 };
  }
  return out;
}

function draftsFor_(token) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_DRAFTS);
  var out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim() !== token) continue;
    var data = {};
    try { data = JSON.parse(rows[i][3] || '{}'); } catch (err) { data = {}; }
    out[String(rows[i][1]).trim()] = { answers: data, updatedAt: toIso_(rows[i][2]) };
  }
  return out;
}

function toIso_(v) {
  if (v instanceof Date) return v.toISOString();
  if (!v) return '';
  return String(v);
}

function apiStudent_(p) {
  var s = findStudent_(String(p.token || '').trim());
  if (!s) return { ok: false, error: 'not_found' };
  return {
    ok: true,
    student: { name: s.name, mentor: s.mentor, createdAt: toIso_(s.createdAt) },
    submitted: answersFor_(s.token),
    drafts: draftsFor_(s.token)
  };
}

function apiSubmit_(p) {
  var token = String(p.token || '').trim();
  var s = findStudent_(token);
  if (!s) return { ok: false, error: 'not_found' };

  var dayId = String(p.dayId || '').trim();
  if (!dayId) return { ok: false, error: 'no_day' };

  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (err) { return { ok: false, error: 'busy' }; }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var an = ss.getSheetByName(SH_ANSWERS);
    var existing = answersFor_(token);
    var now = new Date();
    var answers = p.answers || {};
    var row = [
      token, s.name, dayId, Number(p.n || 0), String(p.module || ''), String(p.title || ''),
      now, String(p.readable || ''), JSON.stringify(answers)
    ];

    if (existing[dayId]) {
      an.getRange(existing[dayId].row, 1, 1, 9).setValues([row]);
    } else {
      an.appendRow(row);
      an.getRange(an.getLastRow(), 7).setNumberFormat('dd.MM.yyyy HH:mm');
      an.getRange(an.getLastRow(), 8).setWrap(true).setVerticalAlignment('top');
      an.getRange(an.getLastRow(), 1, 1, 9).setVerticalAlignment('top');
    }

    clearDraft_(token, dayId);

    var count = Object.keys(answersFor_(token)).length;
    var st = ss.getSheetByName(SH_STUDENTS);
    st.getRange(s.row, 7).setValue(count + ' / ' + (p.total || 31));
    st.getRange(s.row, 8).setValue(Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm'));

    return { ok: true, submittedAt: now.toISOString(), count: count };
  } finally {
    lock.releaseLock();
  }
}

function apiDraft_(p) {
  var token = String(p.token || '').trim();
  var s = findStudent_(token);
  if (!s) return { ok: false, error: 'not_found' };
  var dayId = String(p.dayId || '').trim();
  if (!dayId) return { ok: false, error: 'no_day' };

  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_DRAFTS);
  var now = new Date();
  var payload = JSON.stringify(p.answers || {});
  var found = -1;
  if (sh.getLastRow() > 1) {
    var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === token && String(rows[i][1]).trim() === dayId) { found = i + 2; break; }
    }
  }
  if (found > 0) sh.getRange(found, 1, 1, 4).setValues([[token, dayId, now, payload]]);
  else sh.appendRow([token, dayId, now, payload]);
  return { ok: true };
}

function clearDraft_(token, dayId) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_DRAFTS);
  if (!sh || sh.getLastRow() < 2) return;
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  for (var i = rows.length - 1; i >= 0; i--) {
    if (String(rows[i][0]).trim() === token && String(rows[i][1]).trim() === dayId) sh.deleteRow(i + 2);
  }
}

/* ── Наставник ── */

function findMentor_(key) {
  if (!key) return null;
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_MENTORS);
  if (!sh || sh.getLastRow() < 2) return null;
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][1]).trim() === key) return { name: String(rows[i][0]), key: key };
  }
  return null;
}

function apiMentor_(p) {
  var m = findMentor_(String(p.key || '').trim());
  if (!m) return { ok: false, error: 'not_found' };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var st = ss.getSheetByName(SH_STUDENTS);
  var an = ss.getSheetByName(SH_ANSWERS);

  var byToken = {};
  if (an && an.getLastRow() > 1) {
    var rows = an.getRange(2, 1, an.getLastRow() - 1, 7).getValues();
    for (var i = 0; i < rows.length; i++) {
      var tk = String(rows[i][0]).trim();
      if (!tk) continue;
      if (!byToken[tk]) byToken[tk] = { days: {}, last: null };
      byToken[tk].days[String(rows[i][2]).trim()] = toIso_(rows[i][6]);
      var d = rows[i][6];
      if (d instanceof Date && (!byToken[tk].last || d > byToken[tk].last)) byToken[tk].last = d;
    }
  }

  var students = [];
  if (st && st.getLastRow() > 1) {
    var srows = st.getRange(2, 1, st.getLastRow() - 1, 9).getValues();
    for (var j = 0; j < srows.length; j++) {
      var name = String(srows[j][0]).trim();
      var token = String(srows[j][3]).trim();
      if (!name || !token) continue;
      var info = byToken[token] || { days: {}, last: null };
      students.push({
        name: name,
        contact: String(srows[j][1] || ''),
        mentor: String(srows[j][2] || ''),
        token: token,
        link: String(srows[j][4] || ''),
        createdAt: toIso_(srows[j][5]),
        note: String(srows[j][8] || ''),
        days: info.days,
        lastActivity: info.last ? info.last.toISOString() : ''
      });
    }
  }
  return { ok: true, mentor: { name: m.name }, students: students };
}

function apiMentorStudent_(p) {
  var m = findMentor_(String(p.key || '').trim());
  if (!m) return { ok: false, error: 'not_found' };
  var token = String(p.token || '').trim();
  var s = findStudent_(token);
  if (!s) return { ok: false, error: 'student_not_found' };
  return {
    ok: true,
    student: { name: s.name, contact: s.contact, mentor: s.mentor, token: token },
    submitted: answersFor_(token)
  };
}
