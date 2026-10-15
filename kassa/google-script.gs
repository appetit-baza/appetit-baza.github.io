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
