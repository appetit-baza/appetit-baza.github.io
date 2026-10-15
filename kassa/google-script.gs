/**
 * Кассовый отчёт «Аппетит» — серверная часть в Google Таблице.
 *
 * Что делает:
 *  - проверяет PIN администратора по листу «Админы»;
 *  - принимает отчёты с сайта и раскладывает их по листам
 *    «Отчёты», «Курьеры», «Приход-расход».
 *
 * Как поставить (один раз):
 *  1. Создайте новую Google Таблицу.
 *  2. Расширения → Apps Script. Удалите всё в окне и вставьте этот файл целиком. Сохраните.
 *  3. Сверху выберите функцию setup и нажмите «Выполнить». Разрешите доступ.
 *     В таблице появятся листы с заголовками.
 *  4. Начать развёртывание → Новое развёртывание → тип «Веб-приложение».
 *     Запуск от имени: «Я». У кого есть доступ: «Все».
 *     Нажмите «Начать развёртывание» и скопируйте ссылку веб-приложения.
 *  5. Пришлите ссылку разработчику сайта (её нужно вписать в сайт один раз).
 *
 * Как заводить курьеров: лист «Справочник курьеров», одна строка на курьера.
 *   Филиал | ФИО курьера | Активен. Уволился — поставьте «нет» в «Активен».
 *   Филиал пишите так же, как в списке на сайте.
 *
 * Как заводить администраторов: лист «Админы», одна строка на человека.
 *   ФИО | Филиал | PIN | Активен
 *   PIN — от 4 до 8 цифр, у каждого свой. Чтобы закрыть доступ, поставьте в «Активен» «нет».
 *   Филиал пишите так же, как в списке на сайте (например: Казань, Академика Глушко 35).
 */

// ID таблицы (часть ссылки между /d/ и /edit). Пусто — если скрипт создан из самой таблицы.
var SPREADSHEET_ID = '';
function ss_() { return SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActive(); }

var SHEET_ADMINS = 'Админы';
var SHEET_REPORTS = 'Отчёты';
var SHEET_COURIERS = 'Курьеры';
var SHEET_MONEY = 'Приход-расход';
var SHEET_LOG = 'Журнал входов';
var SHEET_STAFF = 'Справочник курьеров';

var HEADERS = {};
HEADERS[SHEET_ADMINS] = ['ФИО', 'Филиал', 'PIN', 'Активен'];
HEADERS[SHEET_REPORTS] = ['Отправлено', 'Дата смены', 'Филиал', 'Администратор', 'Выручка всего', 'Наличка',
  'Б/Н расчет (картой)', 'Деливери/ЯЕ', 'МП (online)', 'Чиббис', 'Сбермаркет', 'Бонусы iiko (iikocard)',
  'Итого по 041', 'Курьерам всего', 'Делали изъятие', 'Не делали изъятие', 'Прочие расходы',
  'из них вычтено', 'Приход', 'Передана сумма', 'Повторная отправка', 'ID отчёта'];
HEADERS[SHEET_COURIERS] = ['ID отчёта', 'Дата смены', 'Филиал', 'Курьер', 'С', 'По', 'За часы', 'Заказов',
  'За заказы', 'Итого', 'Изъятие', 'Нет в справочнике'];
HEADERS[SHEET_MONEY] = ['ID отчёта', 'Дата смены', 'Филиал', 'Тип', 'Дата', 'Статья', 'Сумма', 'Вычтено из переданной'];
HEADERS[SHEET_LOG] = ['Время', 'Событие', 'ФИО', 'Филиал'];
HEADERS[SHEET_STAFF] = ['Филиал', 'ФИО курьера', 'Активен'];

var MAX_FAILS = 5;           // неверных PIN подряд с одного устройства
var LOCK_SECONDS = 600;      // блокировка этого устройства на 10 минут
var GLOBAL_MAX_FAILS = 30;   // защита от перебора: столько неверных PIN за 10 минут со всех устройств — вход закрывается для всех

function setup() {
  var ss = ss_();
  Object.keys(HEADERS).forEach(function (name) {
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.appendRow(HEADERS[name]);
      sh.getRange(1, 1, 1, HEADERS[name].length).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
  });
  var admins = ss.getSheetByName(SHEET_ADMINS);
  admins.getRange('C:C').setNumberFormat('@'); // PIN как текст, чтобы не терялись нули впереди
  var def = ss.getSheetByName('Лист1') || ss.getSheetByName('Sheet1');
  if (def && def.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(def);
}

function doGet() {
  return json_({ ok: true, service: 'appetit-kassa' });
}

function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad_request' }); }
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var dev = String(body.device || 'unknown').replace(/[^\w-]/g, '').slice(0, 64) || 'unknown';
    var lockState = isLocked_(dev);
    if (lockState) return json_({ ok: false, error: lockState });
    var admin = findAdmin_(body.pin);
    if (!admin) { var st = registerFail_(dev); return json_({ ok: false, error: st || 'bad_pin', left: st ? 0 : undefined }); }
    resetFails_(dev);
    if (body.action === 'login') {
      log_('Вход', admin.name, admin.branch);
      return json_({ ok: true, name: admin.name, branch: admin.branch, couriers: couriers_() });
    }
    if (body.action === 'couriers') return json_({ ok: true, couriers: couriers_() });
    if (body.action === 'revenue') {
      var res = sendRevenue_(body.report || {}, admin);
      log_('Выручка в MAX: ' + (res.ok ? 'отправлено' : res.error), admin.name, (body.report || {}).branch || admin.branch);
      return json_(res);
    }
    if (body.action === 'history') return json_({ ok: true, shifts: history_(admin) });
    if (body.action === 'shift') {
      var sh = shift_(String(body.id || ''), admin);
      return json_(sh ? { ok: true, shift: sh } : { ok: false, error: 'not_found' });
    }
    if (body.action === 'report') {
      var id = saveReport_(body.report || {}, admin);
      return json_({ ok: true, id: id });
    }
    return json_({ ok: false, error: 'unknown_action' });
  } finally {
    lock.releaseLock();
  }
}

function findAdmin_(pin) {
  pin = String(pin || '').replace(/\D/g, '');
  if (pin.length < 4) return null;
  var sh = ss_().getSheetByName(SHEET_ADMINS);
  var rows = sh.getDataRange().getDisplayValues();
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    var active = String(r[3] || '').trim().toLowerCase();
    if (String(r[2]).replace(/\D/g, '') === pin && active !== 'нет' && active !== 'no' && r[0]) {
      return { name: String(r[0]).trim(), branch: String(r[1] || '').trim() };
    }
  }
  return null;
}

function couriers_() {
  var sh = ss_().getSheetByName(SHEET_STAFF);
  var out = {};
  if (!sh) return out;
  var rows = sh.getDataRange().getDisplayValues();
  for (var i = 1; i < rows.length; i++) {
    var b = String(rows[i][0] || '').trim(), n = String(rows[i][1] || '').trim();
    var active = String(rows[i][2] || '').trim().toLowerCase();
    if (!b || !n || active === 'нет' || active === 'no') continue;
    (out[b] = out[b] || []).push(n);
  }
  Object.keys(out).forEach(function (b) { out[b].sort(function (x, y) { return x.localeCompare(y, 'ru'); }); });
  return out;
}

// ---------- MAX: сообщение о выручке в чат «АДМИНЫ АППЕТИТ»
// Ключ бота хранится в свойствах скрипта (Настройки проекта → Свойства скрипта): MAX_TOKEN.
// ID чата скрипт находит сам по названию и запоминает в MAX_CHAT_ID.
var MAX_API = 'https://platform-api2.max.ru';
var MAX_CHAT_TITLE = 'АДМИНЫ АППЕТИТ';

function maxFetch_(method, path, payload) {
  var token = PropertiesService.getScriptProperties().getProperty('MAX_TOKEN');
  if (!token) return { code: 0, error: 'no_token' };
  var opt = { method: method, headers: { Authorization: token }, muteHttpExceptions: true };
  if (payload) { opt.contentType = 'application/json'; opt.payload = JSON.stringify(payload); }
  var r = UrlFetchApp.fetch(MAX_API + path, opt);
  var body = {};
  try { body = JSON.parse(r.getContentText()); } catch (e) {}
  return { code: r.getResponseCode(), body: body };
}

function maxChatId_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('MAX_CHAT_ID');
  if (id) return id;
  var r = maxFetch_('get', '/chats?count=100');
  var chats = (r.body && r.body.chats) || [];
  for (var i = 0; i < chats.length; i++) {
    if (String(chats[i].title || '').toUpperCase().indexOf(MAX_CHAT_TITLE) >= 0) {
      id = String(chats[i].chat_id);
      props.setProperty('MAX_CHAT_ID', id);
      return id;
    }
  }
  return '';
}

function fmtRub_(n) {
  return Utilities.formatString('%s ₽', String(Math.round(n * 100) / 100).replace(/\B(?=(\d{3})+(?!\d))/g, ' ').replace('.', ','));
}

function sendRevenue_(r, admin) {
  var rev = r.rev || {};
  var names = [['r_cash', 'Наличка'], ['r_card', 'Б/Н картой'], ['r_ye', 'Деливери/ЯЕ'], ['r_mp', 'МП (online)'],
    ['r_chib', 'Чиббис'], ['r_sber', 'Сбермаркет'], ['r_bon', 'Бонусы iiko']];
  var total = 0, lines = [];
  names.forEach(function (x) { var v = num_(rev[x[0]]); total += v; if (v) lines.push(x[1] + ': ' + fmtRub_(v)); });
  if (!total) return { ok: false, error: 'empty' };
  var text = '💰 Выручка ' + (r.dateRu || '') + '\n' + (r.branch || admin.branch) + '\nАдминистратор: ' + admin.name +
    '\n\n' + lines.join('\n') + '\n\nВСЕГО: ' + fmtRub_(total);
  var chat = maxChatId_();
  if (!chat) return { ok: false, error: PropertiesService.getScriptProperties().getProperty('MAX_TOKEN') ? 'no_chat' : 'no_token' };
  var res = maxFetch_('post', '/messages?chat_id=' + encodeURIComponent(chat), { text: text });
  if (res.code >= 200 && res.code < 300) return { ok: true };
  return { ok: false, error: 'max_' + res.code };
}

// Запустите вручную после того, как добавите бота в чат: проверит ключ и найдёт чат.
function testMax() {
  PropertiesService.getScriptProperties().deleteProperty('MAX_CHAT_ID');
  var me = maxFetch_('get', '/me');
  Logger.log('Бот: ' + JSON.stringify(me));
  Logger.log('Чат: ' + maxChatId_());
}

// ---------- Архив: каждый администратор видит только свои отчёты
function d_(v) {
  if (!(v instanceof Date)) return String(v);
  var t = Utilities.formatDate(v, 'Europe/Moscow', 'HH:mm');
  return Utilities.formatDate(v, 'Europe/Moscow', t === '00:00' ? 'dd.MM.yyyy' : 'dd.MM.yyyy HH:mm');
}
function history_(admin) {
  var rows = ss_().getSheetByName(SHEET_REPORTS).getDataRange().getValues();
  var out = [];
  for (var i = rows.length - 1; i >= 1 && out.length < 100; i--) {
    var r = rows[i];
    if (String(r[3]).trim() !== admin.name) continue;
    out.push({ id: String(r[21]), sent: d_(r[0]),
      date: d_(r[1]), branch: String(r[2]), rev: Number(r[4]) || 0, cash: Number(r[5]) || 0,
      ct: Number(r[13]) || 0, ex: Number(r[16]) || 0, pass: Number(r[19]) || 0, repeat: String(r[20]) === 'да' });
  }
  return out;
}

function shift_(id, admin) {
  if (!id) return null;
  var ss = ss_();
  var rows = ss.getSheetByName(SHEET_REPORTS).getDataRange().getValues();
  var head = rows[0], row = null;
  for (var i = 1; i < rows.length; i++) if (String(rows[i][21]) === id) { row = rows[i]; break; }
  if (!row || String(row[3]).trim() !== admin.name) return null;
  var summary = [];
  for (var j = 1; j < head.length - 1; j++) {
    var v = row[j];
    if (v === '' || v === null || (v === 0 && String(head[j]) !== 'Передана сумма')) continue;
    summary.push([String(head[j]), v instanceof Date ? d_(v) : v]);
  }
  var cour = ss.getSheetByName(SHEET_COURIERS).getDataRange().getDisplayValues().filter(function (r) { return r[0] === id; })
    .map(function (r) { return { name: r[3], from: r[4], to: r[5], hours: r[6], qty: r[7], orders: r[8], total: r[9], w: r[10], other: r[11] }; });
  var money = ss.getSheetByName(SHEET_MONEY).getDataRange().getDisplayValues().filter(function (r) { return r[0] === id; })
    .map(function (r) { return { type: r[3], date: r[4], art: r[5], sum: r[6], w: r[7] }; });
  return { id: id, summary: summary, cour: cour, money: money };
}

function num_(v) {
  var n = parseFloat(String(v === undefined || v === null ? '' : v).replace(/\s/g, '').replace(',', '.'));
  return isFinite(n) ? n : 0;
}

function saveReport_(r, admin) {
  var ss = ss_();
  var id = Utilities.formatDate(new Date(), 'Europe/Moscow', 'yyMMdd-HHmmss') + '-' + Math.floor(Math.random() * 900 + 100);
  var t = r.totals || {};
  var rev = r.rev || {};
  var branch = r.branch || admin.branch;
  var rep = ss.getSheetByName(SHEET_REPORTS);
  var repeat = rep.getDataRange().getDisplayValues().some(function (row, i) {
    return i > 0 && row[1] === r.dateRu && row[2] === branch;
  });
  rep.appendRow([new Date(), r.dateRu, branch, admin.name, num_(t.rev), num_(rev.r_cash), num_(rev.r_card),
    num_(rev.r_ye), num_(rev.r_mp), num_(rev.r_chib), num_(rev.r_sber), num_(rev.r_bon),
    r.r041 ? num_(r.r041) : '', num_(t.ct), num_(t.cw), num_(t.cn), num_(t.ex), num_(t.exW), num_(t.inc),
    num_(t.pass), repeat ? 'да' : '', id]);
  var cs = ss.getSheetByName(SHEET_COURIERS);
  (r.cour || []).forEach(function (c) {
    var h = num_(c.hours), o = num_(c.orders);
    if (!h && !o && !c.name) return;
    cs.appendRow([id, r.dateRu, branch, c.name, c.from, c.to, h, num_(c.qty), o, h + o, c.w ? 'да' : 'нет (сейф)', c.other ? 'да' : '']);
  });
  var ms = ss.getSheetByName(SHEET_MONEY);
  (r.exp || []).forEach(function (x) {
    if (!num_(x.sum)) return;
    ms.appendRow([id, r.dateRu, branch, 'Расход', x.dateRu, x.art, num_(x.sum), x.w ? 'да' : '']);
  });
  (r.inc || []).forEach(function (x) {
    if (!num_(x.sum)) return;
    ms.appendRow([id, r.dateRu, branch, 'Приход', x.dateRu, x.art, num_(x.sum), '']);
  });
  log_('Отчёт ' + id, admin.name, branch);
  return id;
}

function isLocked_(dev) {
  var c = CacheService.getScriptCache();
  if (c.get('lockAll')) return 'locked_all';
  if (c.get('lock_' + dev)) return 'locked';
  return '';
}
function registerFail_(dev) {
  var c = CacheService.getScriptCache();
  var n = Number(c.get('fails_' + dev) || 0) + 1;
  c.put('fails_' + dev, String(n), LOCK_SECONDS);
  var g = Number(c.get('failsAll') || 0) + 1;
  c.put('failsAll', String(g), LOCK_SECONDS);
  if (g >= GLOBAL_MAX_FAILS) {
    c.put('lockAll', '1', LOCK_SECONDS);
    c.remove('failsAll');
    log_('Блокировка входа для всех: много неверных PIN', '', '');
    return 'locked_all';
  }
  if (n >= MAX_FAILS) {
    c.put('lock_' + dev, '1', LOCK_SECONDS);
    c.remove('fails_' + dev);
    log_('Блокировка устройства ' + dev.slice(0, 8) + ': ' + MAX_FAILS + ' неверных PIN', '', '');
    return 'locked';
  }
  return '';
}
function resetFails_(dev) { CacheService.getScriptCache().remove('fails_' + dev); }

function log_(event, name, branch) {
  var sh = ss_().getSheetByName(SHEET_LOG);
  if (sh) sh.appendRow([new Date(), event, name, branch]);
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
