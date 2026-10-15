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

var MAX_FAILS = 15;          // неудачных попыток подряд
var LOCK_SECONDS = 600;      // блокировка входа на 10 минут

function setup() {
  var ss = SpreadsheetApp.getActive();
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
    if (isLocked_()) return json_({ ok: false, error: 'locked' });
    var admin = findAdmin_(body.pin);
    if (!admin) { registerFail_(); return json_({ ok: false, error: 'bad_pin' }); }
    resetFails_();
    if (body.action === 'login') {
      log_('Вход', admin.name, admin.branch);
      return json_({ ok: true, name: admin.name, branch: admin.branch, couriers: couriers_() });
    }
    if (body.action === 'couriers') return json_({ ok: true, couriers: couriers_() });
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
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_ADMINS);
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
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_STAFF);
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

function num_(v) {
  var n = parseFloat(String(v === undefined || v === null ? '' : v).replace(/\s/g, '').replace(',', '.'));
  return isFinite(n) ? n : 0;
}

function saveReport_(r, admin) {
  var ss = SpreadsheetApp.getActive();
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

function isLocked_() {
  var until = Number(PropertiesService.getScriptProperties().getProperty('lockUntil') || 0);
  return until > Date.now();
}
function registerFail_() {
  var p = PropertiesService.getScriptProperties();
  var n = Number(p.getProperty('fails') || 0) + 1;
  p.setProperty('fails', String(n));
  if (n >= MAX_FAILS) {
    p.setProperty('lockUntil', String(Date.now() + LOCK_SECONDS * 1000));
    p.setProperty('fails', '0');
    log_('Блокировка входа: много неверных PIN', '', '');
  }
}
function resetFails_() { PropertiesService.getScriptProperties().setProperty('fails', '0'); }

function log_(event, name, branch) {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_LOG);
  if (sh) sh.appendRow([new Date(), event, name, branch]);
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
