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

/** Столбцы листа «Ученики» (1 = A) */
var S_NAME = 1, S_CONTACT = 2, S_MENTOR = 3, S_MKEY = 4, S_TOKEN = 5,
    S_LINK = 6, S_CREATED = 7, S_PROGRESS = 8, S_LAST = 9, S_NOTE = 10;

var STUDENT_HEADERS = [
  'Имя', 'Контакт', 'Наставник', 'Ключ наставника', 'Токен',
  'Ссылка для ученика', 'Создан', 'Прогресс', 'Последняя активность', 'Заметки наставника'
];

/** Столбцы листа «Наставники» */
var M_NAME = 1, M_KEY = 2, M_LINK = 3, M_ROLE = 4;
var MENTOR_HEADERS = ['Имя', 'Ключ', 'Ссылка на панель', 'Роль'];

var ROLE_MENTOR = 'наставник';
var ROLE_ADMIN  = 'администратор';

var ANSWER_HEADERS = [
  'Токен', 'Ученик', 'День', '№', 'Модуль', 'Тема', 'Отправлено', 'Ответы', 'JSON'
];
var DRAFT_HEADERS = ['Токен', 'День', 'Изменён', 'JSON'];

var TOTAL_DAYS = 31;

/* ──────────────────────────── Меню таблицы ─────────────────────────── */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📖 Курс')
    .addItem('① Настроить таблицу', 'setupSheets')
    .addItem('② Указать адрес сайта', 'promptSiteUrl')
    .addSeparator()
    .addItem('🔑 Добавить наставника', 'promptAddMentor')
    .addItem('➕ Добавить ученика', 'promptAddStudent')
    .addSeparator()
    .addItem('📋 Скопировать ссылку выбранного ученика', 'showSelectedLink')
    .addItem('🔗 Выдать ссылки всем без токена', 'fillMissingTokens')
    .addItem('👥 Связать учеников с наставниками', 'linkStudentsToMentors')
    .addItem('🔄 Пересчитать прогресс', 'recalcProgress')
    .addToUi();
}

/* ─────────────────────────── Инициализация ─────────────────────────── */

function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var st = ensureSheet_(ss, SH_STUDENTS, STUDENT_HEADERS);
  var widths = [200, 160, 160, 150, 190, 360, 140, 100, 150, 260];
  for (var i = 0; i < widths.length; i++) st.setColumnWidth(i + 1, widths[i]);
  st.hideColumns(S_MKEY);

  var an = ensureSheet_(ss, SH_ANSWERS, ANSWER_HEADERS);
  var aw = [150, 170, 60, 50, 130, 230, 150, 640, 120];
  for (var j = 0; j < aw.length; j++) an.setColumnWidth(j + 1, aw[j]);
  an.hideColumns(9);

  var dr = ensureSheet_(ss, SH_DRAFTS, DRAFT_HEADERS);
  dr.hideSheet();

  var mn = ensureSheet_(ss, SH_MENTORS, MENTOR_HEADERS);
  mn.setColumnWidth(M_NAME, 200);
  mn.setColumnWidth(M_KEY, 190);
  mn.setColumnWidth(M_LINK, 380);
  mn.setColumnWidth(M_ROLE, 150);

  var cf = ensureSheet_(ss, SH_CONFIG, ['Параметр', 'Значение']);
  if (cf.getLastRow() < 2) {
    cf.getRange(2, 1, 1, 2).setValues([['SITE_URL', '']]);
    cf.getRange(3, 1, 1, 2).setValues([['Подсказка', 'Укажите в SITE_URL адрес сайта, например https://user.github.io/kreshenie/']]);
  }
  cf.setColumnWidth(1, 160);
  cf.setColumnWidth(2, 520);

  refreshValidation_();

  SpreadsheetApp.getUi().alert('Готово',
    'Листы созданы и настроены.\n\nДальше:\n② Указать адрес сайта\n🔑 Добавить наставника\n➕ Добавить ученика',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function ensureSheet_(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  var width = Math.max(headers.length, sh.getLastColumn() || 1);
  var existing = sh.getRange(1, 1, 1, width).getValues()[0];
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

/**
 * Выпадающие списки: наставник у ученика — из имён на листе «Наставники»,
 * роль наставника — из двух допустимых значений.
 * Вызывается при настройке и после добавления наставника.
 */
function refreshValidation_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var st = ss.getSheetByName(SH_STUDENTS);
  var mn = ss.getSheetByName(SH_MENTORS);
  if (!st || !mn) return;

  try {
    var last = Math.max(mn.getLastRow(), 2);
    var namesRange = mn.getRange(2, M_NAME, last - 1, 1);
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(namesRange, true)
      .setAllowInvalid(false)
      .setHelpText('Выберите наставника из списка на листе «Наставники».')
      .build();
    st.getRange(2, S_MENTOR, Math.max(st.getMaxRows() - 1, 1), 1).setDataValidation(rule);

    var roleRule = SpreadsheetApp.newDataValidation()
      .requireValueInList([ROLE_MENTOR, ROLE_ADMIN], true)
      .setAllowInvalid(false)
      .setHelpText('«наставник» видит своих учеников, «администратор» — всех.')
      .build();
    mn.getRange(2, M_ROLE, Math.max(mn.getMaxRows() - 1, 1), 1).setDataValidation(roleRule);
  } catch (err) { /* валидация необязательна */ }
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

/** Все наставники: [{row, name, key, role}] */
function listMentors_() {
  var mn = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_MENTORS);
  var out = [];
  if (!mn || mn.getLastRow() < 2) return out;
  var rows = mn.getRange(2, 1, mn.getLastRow() - 1, 4).getValues();
  for (var i = 0; i < rows.length; i++) {
    var name = String(rows[i][M_NAME - 1] || '').trim();
    var key = String(rows[i][M_KEY - 1] || '').trim();
    if (!name && !key) continue;
    out.push({
      row: i + 2, name: name, key: key,
      role: String(rows[i][M_ROLE - 1] || ROLE_MENTOR).trim().toLowerCase() || ROLE_MENTOR
    });
  }
  return out;
}

function findMentorByKey_(key) {
  if (!key) return null;
  var all = listMentors_();
  for (var i = 0; i < all.length; i++) if (all[i].key === key) return all[i];
  return null;
}

function findMentorByName_(name) {
  var needle = String(name || '').trim().toLowerCase();
  if (!needle) return null;
  var all = listMentors_();
  for (var i = 0; i < all.length; i++) if (all[i].name.toLowerCase() === needle) return all[i];
  return null;
}

function promptAddMentor() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('Новый наставник', 'Имя наставника:', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var name = res.getResponseText().trim();
  if (!name) return;

  if (findMentorByName_(name)) {
    ui.alert('Такой наставник уже есть', 'Имя «' + name + '» уже занято. Имена должны быть разными.', ui.ButtonSet.OK);
    return;
  }

  var roleRes = ui.alert('Роль наставника',
    'Сделать «' + name + '» администратором?\n\n' +
    'Да — будет видеть учеников всех наставников.\n' +
    'Нет — только своих.',
    ui.ButtonSet.YES_NO);
  var role = roleRes === ui.Button.YES ? ROLE_ADMIN : ROLE_MENTOR;

  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_MENTORS);
  var key = makeToken_('m-');
  sh.appendRow([name, key, mentorLink_(key), role]);
  refreshValidation_();

  ui.alert('Наставник создан',
    name + ' — ' + role + '\n\nСсылка на панель:\n' + mentorLink_(key) +
    '\n\nЭту ссылку нельзя давать ученикам.',
    ui.ButtonSet.OK);
}

function promptAddStudent() {
  var ui = SpreadsheetApp.getUi();
  var mentors = listMentors_();
  if (!mentors.length) {
    ui.alert('Сначала добавьте наставника',
      'Меню «📖 Курс → 🔑 Добавить наставника».', ui.ButtonSet.OK);
    return;
  }

  var res = ui.prompt('Новый ученик', 'Имя и фамилия ученика:', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var name = res.getResponseText().trim();
  if (!name) return;

  var mentor = mentors[0];
  if (mentors.length > 1) {
    var list = mentors.map(function (m, i) { return (i + 1) + '. ' + m.name; }).join('\n');
    var mr = ui.prompt('Наставник ученика',
      'Кто ведёт «' + name + '»? Введите номер:\n\n' + list, ui.ButtonSet.OK_CANCEL);
    if (mr.getSelectedButton() !== ui.Button.OK) return;
    var idx = parseInt(mr.getResponseText().trim(), 10);
    if (!idx || idx < 1 || idx > mentors.length) {
      ui.alert('Не понял номер', 'Ученик не создан. Попробуйте ещё раз.', ui.ButtonSet.OK);
      return;
    }
    mentor = mentors[idx - 1];
  }

  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_STUDENTS);
  var token = makeToken_('s-');
  sh.appendRow([name, '', mentor.name, mentor.key, token, studentLink_(token),
                new Date(), '0 / ' + TOTAL_DAYS, '', '']);
  sh.getRange(sh.getLastRow(), S_CREATED).setNumberFormat('dd.MM.yyyy HH:mm');

  ui.alert('Ученик создан',
    name + '\nНаставник: ' + mentor.name + '\n\nИндивидуальная ссылка:\n' + studentLink_(token) +
    '\n\nСсылка также записана в таблицу.',
    ui.ButtonSet.OK);
}

/**
 * Простой триггер. Реагирует на два столбца листа «Ученики»:
 *  A «Имя»       → выдаёт токен и ссылку;
 *  C «Наставник» → подставляет его ключ в скрытый столбец D.
 */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    var sh = e.range.getSheet();
    if (sh.getName() !== SH_STUDENTS) return;
    var row = e.range.getRow();
    if (row < 2) return;
    var col = e.range.getColumn();

    if (col === S_MENTOR) {
      var m = findMentorByName_(sh.getRange(row, S_MENTOR).getValue());
      sh.getRange(row, S_MKEY).setValue(m ? m.key : '');
      return;
    }

    if (col !== S_NAME) return;
    var name = String(sh.getRange(row, S_NAME).getValue()).trim();
    if (!name) return;
    if (String(sh.getRange(row, S_TOKEN).getValue()).trim()) return; // токен уже есть

    var token = makeToken_('s-');
    sh.getRange(row, S_TOKEN).setValue(token);
    sh.getRange(row, S_LINK).setValue(studentLink_(token));
    sh.getRange(row, S_CREATED).setValue(new Date()).setNumberFormat('dd.MM.yyyy HH:mm');
    sh.getRange(row, S_PROGRESS).setValue('0 / ' + TOTAL_DAYS);

    // если наставник уже вписан — сразу проставим его ключ
    var m2 = findMentorByName_(sh.getRange(row, S_MENTOR).getValue());
    if (m2) sh.getRange(row, S_MKEY).setValue(m2.key);
  } catch (err) { /* тихо */ }
}

function fillMissingTokens() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_STUDENTS);
  var last = sh.getLastRow();
  var added = 0;
  for (var r = 2; r <= last; r++) {
    var name = String(sh.getRange(r, S_NAME).getValue()).trim();
    if (!name) continue;
    if (String(sh.getRange(r, S_TOKEN).getValue()).trim()) continue;
    var token = makeToken_('s-');
    sh.getRange(r, S_TOKEN).setValue(token);
    sh.getRange(r, S_LINK).setValue(studentLink_(token));
    if (!sh.getRange(r, S_CREATED).getValue()) {
      sh.getRange(r, S_CREATED).setValue(new Date()).setNumberFormat('dd.MM.yyyy HH:mm');
    }
    added++;
  }
  refreshAllLinks_();
  SpreadsheetApp.getUi().alert('Готово', 'Выдано новых ссылок: ' + added, SpreadsheetApp.getUi().ButtonSet.OK);
}

/** Проставляет ключ наставника всем строкам, где он пуст, по имени в столбце C. */
function linkStudentsToMentors() {
  var ui = SpreadsheetApp.getUi();
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_STUDENTS);
  refreshValidation_();
  if (!sh || sh.getLastRow() < 2) { ui.alert('Учеников пока нет.'); return; }

  var n = sh.getLastRow() - 1;
  var names = sh.getRange(2, S_MENTOR, n, 1).getValues();
  var keys  = sh.getRange(2, S_MKEY, n, 1).getValues();
  var linked = 0, unknown = [], empty = 0;

  for (var i = 0; i < n; i++) {
    var nm = String(names[i][0] || '').trim();
    if (!nm) { keys[i][0] = ''; empty++; continue; }
    var m = findMentorByName_(nm);
    if (!m) { unknown.push(nm); continue; }
    if (keys[i][0] !== m.key) { keys[i][0] = m.key; linked++; }
  }
  sh.getRange(2, S_MKEY, n, 1).setValues(keys);

  var msg = 'Связано записей: ' + linked;
  if (empty) msg += '\nБез наставника: ' + empty + ' — их видит только администратор.';
  if (unknown.length) {
    var uniq = unknown.filter(function (v, i, a) { return a.indexOf(v) === i; });
    msg += '\n\nНе нашёл таких наставников: ' + uniq.join(', ') +
           '\nДобавьте их на лист «Наставники» и повторите.';
  }
  ui.alert('Готово', msg, ui.ButtonSet.OK);
}

function refreshAllLinks_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var st = ss.getSheetByName(SH_STUDENTS);
  if (st && st.getLastRow() > 1) {
    var n = st.getLastRow() - 1;
    var tokens = st.getRange(2, S_TOKEN, n, 1).getValues();
    st.getRange(2, S_LINK, n, 1).setValues(tokens.map(function (t) {
      var v = String(t[0]).trim();
      return [v ? studentLink_(v) : ''];
    }));
  }
  var mn = ss.getSheetByName(SH_MENTORS);
  if (mn && mn.getLastRow() > 1) {
    var m = mn.getLastRow() - 1;
    var keys = mn.getRange(2, M_KEY, m, 1).getValues();
    mn.getRange(2, M_LINK, m, 1).setValues(keys.map(function (t) {
      var v = String(t[0]).trim();
      return [v ? mentorLink_(v) : ''];
    }));
  }
}

function showSelectedLink() {
  var ui = SpreadsheetApp.getUi();
  var sh = SpreadsheetApp.getActiveSheet();
  if (sh.getName() !== SH_STUDENTS) { ui.alert('Откройте лист «Ученики» и выберите строку ученика.'); return; }
  var row = sh.getActiveRange().getRow();
  if (row < 2) { ui.alert('Выберите строку с учеником.'); return; }
  ui.alert('Ссылка ученика',
    sh.getRange(row, S_NAME).getValue() + '\n\n' + sh.getRange(row, S_LINK).getValue(),
    ui.ButtonSet.OK);
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
  var tokens = st.getRange(2, S_TOKEN, n, 1).getValues();
  st.getRange(2, S_PROGRESS, n, 2).setValues(tokens.map(function (t) {
    var tk = String(t[0]).trim();
    var la = lastAt[tk] ? Utilities.formatDate(lastAt[tk], Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm') : '';
    return [(counts[tk] || 0) + ' / ' + TOTAL_DAYS, la];
  }));
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
    if (action === 'ping')          return json_({ ok: true, version: 2 });
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
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, STUDENT_HEADERS.length).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][S_TOKEN - 1]).trim() === token) {
      return {
        row: i + 2,
        name: String(rows[i][S_NAME - 1]),
        contact: String(rows[i][S_CONTACT - 1]),
        mentor: String(rows[i][S_MENTOR - 1]),
        mentorKey: String(rows[i][S_MKEY - 1]).trim(),
        token: token,
        createdAt: rows[i][S_CREATED - 1]
      };
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
    var data = {};
    try { data = JSON.parse(rows[i][8] || '{}'); } catch (err) { data = {}; }
    out[String(rows[i][2]).trim()] = { answers: data, submittedAt: toIso_(rows[i][6]), row: i + 2 };
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
    var row = [
      token, s.name, dayId, Number(p.n || 0), String(p.module || ''), String(p.title || ''),
      now, String(p.readable || ''), JSON.stringify(p.answers || {})
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
    st.getRange(s.row, S_PROGRESS).setValue(count + ' / ' + (p.total || TOTAL_DAYS));
    st.getRange(s.row, S_LAST).setValue(Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm'));

    return { ok: true, submittedAt: now.toISOString(), count: count };
  } finally {
    lock.releaseLock();
  }
}

function apiDraft_(p) {
  var token = String(p.token || '').trim();
  if (!findStudent_(token)) return { ok: false, error: 'not_found' };
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

/** true, если наставник вправе видеть этого ученика. */
function mentorSees_(mentor, studentMentorKey) {
  if (mentor.role === ROLE_ADMIN) return true;
  return !!studentMentorKey && studentMentorKey === mentor.key;
}

function apiMentor_(p) {
  var m = findMentorByKey_(String(p.key || '').trim());
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

  var students = [], unassigned = 0;
  if (st && st.getLastRow() > 1) {
    var srows = st.getRange(2, 1, st.getLastRow() - 1, STUDENT_HEADERS.length).getValues();
    for (var j = 0; j < srows.length; j++) {
      var name = String(srows[j][S_NAME - 1]).trim();
      var token = String(srows[j][S_TOKEN - 1]).trim();
      if (!name || !token) continue;

      var mkey = String(srows[j][S_MKEY - 1]).trim();
      if (!mkey) unassigned++;
      if (!mentorSees_(m, mkey)) continue;

      var info = byToken[token] || { days: {}, last: null };
      students.push({
        name: name,
        contact: String(srows[j][S_CONTACT - 1] || ''),
        mentor: String(srows[j][S_MENTOR - 1] || ''),
        mentorKey: mkey,
        token: token,
        link: String(srows[j][S_LINK - 1] || ''),
        createdAt: toIso_(srows[j][S_CREATED - 1]),
        note: String(srows[j][S_NOTE - 1] || ''),
        days: info.days,
        lastActivity: info.last ? info.last.toISOString() : ''
      });
    }
  }

  return {
    ok: true,
    mentor: { name: m.name, role: m.role, key: m.key, isAdmin: m.role === ROLE_ADMIN },
    unassigned: m.role === ROLE_ADMIN ? unassigned : 0,
    students: students
  };
}

function apiMentorStudent_(p) {
  var m = findMentorByKey_(String(p.key || '').trim());
  if (!m) return { ok: false, error: 'not_found' };

  var token = String(p.token || '').trim();
  var s = findStudent_(token);
  if (!s) return { ok: false, error: 'student_not_found' };
  if (!mentorSees_(m, s.mentorKey)) return { ok: false, error: 'forbidden' };

  return {
    ok: true,
    student: { name: s.name, contact: s.contact, mentor: s.mentor, token: token },
    submitted: answersFor_(token)
  };
}
