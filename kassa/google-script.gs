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
 *   ФИО | Филиал | PIN | Активен | Роль
 *   PIN — от 4 до 8 цифр, у каждого свой. Чтобы закрыть доступ, поставьте в «Активен» «нет».
 *   Роль: пусто — обычный администратор (видит только свои смены);
 *   «руководитель» — видит на сайте все отчёты всех филиалов.
 *   «просмотр» — видит всё то же, что руководитель, но ничего не может менять.
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
var SHEET_FIX = 'Исправления';
var SHEET_DRAFTS = 'Черновики';
var SHEET_EXTRA = 'Доплаты курьерам';
var SHEET_CASH = 'Остатки касс';

var HEADERS = {};
HEADERS[SHEET_ADMINS] = ['ФИО', 'Филиал', 'PIN', 'Активен', 'Роль'];
HEADERS[SHEET_REPORTS] = ['Отправлено', 'Дата смены', 'Филиал', 'Администратор', 'Выручка всего', 'Наличные',
  'Карта', 'Яндекс/Деливери', 'Онлайн', 'Чиббис', 'Сбермаркет', 'Бонусы iiko (iikocard)',
  'Итого по 041', 'Курьерам всего', 'Делали изъятие', 'Не делали изъятие', 'Прочие расходы',
  'из них вычтено', 'Приход', 'Передана сумма', 'Повторная отправка', 'ID отчёта', 'Заменён отчётом', 'Безнал факт (по сверкам)'];
HEADERS[SHEET_COURIERS] = ['ID отчёта', 'Дата смены', 'Филиал', 'Курьер', 'С', 'По', 'За часы', 'Заказов',
  'За заказы', 'Итого', 'Изъятие', 'Нет в справочнике', 'По зонам'];
HEADERS[SHEET_MONEY] = ['ID отчёта', 'Дата смены', 'Филиал', 'Тип', 'Дата', 'Статья', 'Сумма', 'Вычтено из переданной'];
HEADERS[SHEET_LOG] = ['Время', 'Событие', 'ФИО', 'Филиал'];
HEADERS[SHEET_STAFF] = ['Филиал', 'ФИО курьера', 'Активен'];
HEADERS[SHEET_DRAFTS] = ['Администратор', 'Обновлён', 'Черновик'];
HEADERS[SHEET_EXTRA] = ['ID отчёта', 'Дата смены', 'Филиал', 'Администратор', 'Курьер', 'Вид', 'Зона',
  'Причина / комментарий', 'Кол-во', 'Сумма'];
HEADERS[SHEET_FIX] = ['Запрошено', 'ID отчёта', 'Дата смены', 'Филиал', 'Администратор', 'Причина',
  'Статус', 'Решение', 'Новый отчёт', 'ID запроса'];

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
  addRoleColumn();
}

// Добавляет столбец «Роль» на лист «Админы» (для уже работающей таблицы — запустить один раз).
function addRoleColumn() {
  var sh = ss_().getSheetByName(SHEET_ADMINS);
  if (!String(sh.getRange('E1').getValue()).trim()) sh.getRange('E1').setValue('Роль').setFontWeight('bold');
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(['руководитель', 'просмотр'], true).setAllowInvalid(true).build();
  sh.getRange('E2:E').setDataValidation(rule);
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
      return json_({ ok: true, name: admin.name, branch: admin.branch, boss: admin.boss, viewer: !!admin.viewer, couriers: couriers_(),
        fixPending: admin.boss ? fixList_().length : 0, draft: admin.boss ? null : draftGet_(admin) });
    }
    if (body.action === 'couriers') return json_({ ok: true, couriers: couriers_() });
    if (body.action === 'revenue') {
      var res = sendRevenue_(body.report || {}, admin);
      log_('Выручка в MAX: ' + (res.ok ? 'отправлено' : res.error), admin.name, (body.report || {}).branch || admin.branch);
      return json_(res);
    }
    if (admin.viewer && /^(revenue|draft_save|fix_request|fix_decide|report|cash_start)$/.test(String(body.action))) return json_({ ok: false, error: 'view_only' });
    if (body.action === 'cash_get') return json_(admin.boss ? cashGet_(admin) : { ok: false, error: 'not_boss' });
    if (body.action === 'cash_start') return json_(admin.boss ? cashStart_(body, admin) : { ok: false, error: 'not_boss' });
    if (body.action === 'history') return json_({ ok: true, boss: admin.boss, viewer: !!admin.viewer, shifts: history_(admin, body),
      fixPending: admin.boss ? fixList_().length : 0 });
    if (body.action === 'draft_save') return json_(draftSave_(admin, body.draft || null));
    if (body.action === 'draft_get') return json_({ ok: true, draft: draftGet_(admin) });
    if (body.action === 'fix_request') return json_(fixRequest_(String(body.id || ''), String(body.reason || ''), admin));
    if (body.action === 'fix_list') return json_(admin.boss ? { ok: true, list: fixList_() } : { ok: false, error: 'not_boss' });
    if (body.action === 'fix_decide') return json_(admin.boss ? fixDecide_(String(body.rid || ''), !!body.approve, admin) : { ok: false, error: 'not_boss' });
    if (body.action === 'shift') {
      var sh = shift_(String(body.id || ''), admin);
      return json_(sh ? { ok: true, shift: sh } : { ok: false, error: 'not_found' });
    }
    if (body.action === 'report') {
      var rep = body.report || {}, fix = null;
      if (body.fixOf) {
        fix = fixCheck_(String(body.fixOf), admin);
        if (!fix) return json_({ ok: false, error: 'fix_not_allowed' });
      } else if (!admin.boss && dupExists_(rep, admin)) {
        return json_({ ok: false, error: 'duplicate' });
      }
      var id = saveReport_(rep, admin);
      if (fix) fixDone_(fix, id, admin);
      draftSave_(admin, null);
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
      return { name: String(r[0]).trim(), branch: String(r[1] || '').trim(),
        boss: /руковод|владел|управля|просмотр|наблюд/i.test(String(r[4] || '')),
        // «просмотр» — видит всё, как руководитель, но ничего не меняет (не решает исправления, не отправляет отчёты)
        viewer: /просмотр|наблюд/i.test(String(r[4] || '')) };
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
// Посредник в Яндекс Облаке (функция max-relay): Google → посредник → MAX
var MAX_RELAY = 'https://functions.yandexcloud.net/d4ea85u3v2tqfi0e5mc6';

function maxFetch_(method, path, payload) {
  var token = PropertiesService.getScriptProperties().getProperty('MAX_TOKEN');
  if (!token) return { code: 0, error: 'no_token' };
  // validateHttpsCertificates:false пробовали 22.10.2026 — не помогает (Google всё равно даёт «Ошибка SSL»), проверка оставлена включённой.
  var opt = { method: method, headers: { Authorization: token }, muteHttpExceptions: true };
  if (payload) { opt.contentType = 'application/json'; opt.payload = JSON.stringify(payload); }
  if (MAX_RELAY) {
    // напрямую Google до MAX не достучаться (ошибка SSL) — идём через посредника в Яндекс Облаке
    var rr = UrlFetchApp.fetch(MAX_RELAY, { method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'X-Max-Auth': token }, payload: JSON.stringify({ m: method.toUpperCase(), p: path, b: payload || null }) });
    var out = {};
    try { out = JSON.parse(rr.getContentText()); } catch (e) { return { code: rr.getResponseCode(), body: {}, error: rr.getContentText().slice(0, 200) }; }
    var ob = {};
    try { ob = JSON.parse(out.body); } catch (e) {}
    if (!out.code) return { code: rr.getResponseCode(), body: {}, error: rr.getContentText().slice(0, 300) };
    return { code: out.code, body: ob };
  }
  var r = UrlFetchApp.fetch(MAX_API + path, opt);
  var body = {};
  try { body = JSON.parse(r.getContentText()); } catch (e) {}
  return { code: r.getResponseCode(), body: body };
}

function maxChatId_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('MAX_CHAT_ID');
  if (id) return id;
  // Список чатов (GET /chats) MAX убрал в июне 2026: ID чата берём из последних событий бота
  // (бота добавили в чат или кто-то написал в чат) и проверяем название через GET /chats/{id}.
  var r = maxFetch_('get', '/updates?limit=100&types=bot_added,message_created,user_added,chat_title_changed');
  var ups = (r.body && r.body.updates) || [], seen = {};
  for (var i = 0; i < ups.length; i++) {
    var u = ups[i], c = u.chat_id || (u.message && u.message.recipient && u.message.recipient.chat_id);
    if (!c || seen[c]) continue;
    seen[c] = 1;
    var info = maxFetch_('get', '/chats/' + c);
    var title = String((info.body && info.body.title) || '');
    if (title.toUpperCase().indexOf(MAX_CHAT_TITLE) >= 0) {
      id = String(c);
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
  var names = [['r_mp', 'Онлайн'], ['r_cash', 'Наличные'], ['r_card', 'Карта'], ['r_ye', 'Яндекс/Деливери'],
    ['r_chib', 'Чиббис'], ['r_sber', 'Сбермаркет'], ['r_bon', 'Бонусы iiko']];
  var total = 0, lines = [];
  names.forEach(function (x) { var v = num_(rev[x[0]]); total += v; if (v) lines.push(x[1] + ': ' + fmtRub_(v)); });
  if (!total) return { ok: false, error: 'empty' };
  // филиал — первой строкой, жирным и заглавными, чтобы сразу было видно (формат HTML в MAX)
  var h = function (s) { return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); };
  var text = '📍 <b>' + h(r.branch || admin.branch).toUpperCase() + '</b>\n' +
    '💰 Выручка ' + h(r.dateRu) + '\nАдминистратор: ' + h(admin.name) +
    '\n\n' + h(lines.join('\n')) + '\n\n<b>ВСЕГО: ' + fmtRub_(total) + '</b>\n\n' + (ordText_(r.ord) ? h(ordText_(r.ord)) + '\n\n' : '') + h(redoText_(r.redo)) + '\n❌ Отказы: ' + num_(r.otk);
  var chat = maxChatId_();
  if (!chat) return { ok: false, error: PropertiesService.getScriptProperties().getProperty('MAX_TOKEN') ? 'no_chat' : 'no_token' };
  var res = maxFetch_('post', '/messages?chat_id=' + encodeURIComponent(chat), { text: text, format: 'html' });
  if (res.code >= 200 && res.code < 300) return { ok: true };
  return { ok: false, error: 'max_' + res.code };
}

// Довозы за смену: «🚚 Довозы: 0» или «🚚 Довозы: 3 — Кухня недокомплект 2, Ошибка КЦ 1»
function redoText_(list) {
  var n = 0, parts = [];
  (list || []).forEach(function (x) { var k = num_(x[1]); if (!k) return; n += k; parts.push(String(x[0] || 'причина не указана') + ' ' + k); });
  return '🚚 Довозы: ' + n + (n ? ' — ' + parts.join(', ') : '');
}

// Пробное сообщение в чат — с пометкой «ТЕСТ». Запускать вручную из редактора.
function testMaxSend() {
  var chat = maxChatId_();
  Logger.log('Чат: ' + chat);
  if (!chat) return;
  var r = maxFetch_('post', '/messages?chat_id=' + encodeURIComponent(chat), { text: '🔧 ТЕСТ: проверка связи сайта кассовых отчётов с чатом. Отвечать не нужно.' });
  Logger.log('Ответ MAX: ' + r.code + ' ' + JSON.stringify(r.body).slice(0, 300));
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
// Дата смены → число ГГГГММДД для фильтра по периоду
function dkey_(v) {
  if (v instanceof Date) return Number(Utilities.formatDate(v, 'Europe/Moscow', 'yyyyMMdd'));
  var m = String(v).match(/(\d{1,2})\.(\d{1,2})\.(\d{4})/);
  return m ? Number(m[3]) * 10000 + Number(m[2]) * 100 + Number(m[1]) : 0;
}
function history_(admin, q) {
  q = q || {};
  var rows = ss_().getSheetByName(SHEET_REPORTS).getDataRange().getValues();
  var from = Number(String(q.from || '').replace(/\D/g, '')) || 0;
  var to = Number(String(q.to || '').replace(/\D/g, '')) || 99999999;
  var branch = String(q.branch || '');
  var limit = admin.boss ? 2000 : 100;
  var fx = fixMap_(), rd = redoMap_(), mm = admin.boss ? moneyMap_() : {}, cm = admin.boss ? courMap_() : {};
  var out = [];
  for (var i = rows.length - 1; i >= 1 && out.length < limit; i--) {
    var r = rows[i];
    if (!admin.boss && String(r[3]).trim() !== admin.name) continue;
    if (admin.boss) {
      var k = dkey_(r[1]);
      if (k < from || k > to) continue;
      if (branch && String(r[2]) !== branch) continue;
    }
    out.push({ id: String(r[21]), sent: d_(r[0]), dk: dkey_(r[1]),
      date: d_(r[1]), branch: String(r[2]), admin: String(r[3]), rev: Number(r[4]) || 0, cash: Number(r[5]) || 0,
      card: Number(r[6]) || 0, ye: Number(r[7]) || 0, mp: Number(r[8]) || 0, chib: Number(r[9]) || 0,
      sber: Number(r[10]) || 0, bon: Number(r[11]) || 0,
      ct: Number(r[13]) || 0, ex: Number(r[16]) || 0, inc: Number(r[18]) || 0, pass: Number(r[19]) || 0,
      repeat: String(r[20]) === 'да', fixed: String(r[20]) === 'исправление',
      bfact: r[23] === '' || r[23] === undefined || r[23] === null ? null : Number(r[23]) || 0,
      bdiff: r[23] === '' || r[23] === undefined || r[23] === null ? 0 : Math.round(((Number(r[23]) || 0) - (Number(r[6]) || 0)) * 100) / 100, replaced: String(r[22] || ''), fix: (fx[String(r[21])] || {}).status || '', redo: rd[String(r[21])] || null, money: mm[String(r[21])] || [],
      sentMs: r[0] instanceof Date ? r[0].getTime() : 0, cw: Number(r[14]) || 0, cn: Number(r[15]) || 0, cour: cm[String(r[21])] || [] });
  }
  return out;
}

function shift_(id, admin) {
  if (!id) return null;
  var ss = ss_();
  var rows = ss.getSheetByName(SHEET_REPORTS).getDataRange().getValues();
  var head = rows[0], row = null;
  for (var i = 1; i < rows.length; i++) if (String(rows[i][21]) === id) { row = rows[i]; break; }
  if (!row || (!admin.boss && String(row[3]).trim() !== admin.name)) return null;
  var summary = [];
  for (var j = 1; j <= 20; j++) {
    var v = row[j];
    if (v === '' || v === null || (v === 0 && String(head[j]) !== 'Передана сумма')) continue;
    summary.push([String(head[j]), v instanceof Date ? d_(v) : v]);
  }
  if (row[23] !== '' && row[23] !== null && row[23] !== undefined) summary.push(['Безнал факт (по сверкам)', row[23]]);
  var oq = function (i) { return row[i] === '' || row[i] === undefined || row[i] === null ? '' : String(row[i]); };
  if (oq(24) !== '' || oq(26) !== '') {
    summary.push(['Самовывоз', (oq(24) || 0) + ' / ' + fmtRub_(num_(row[25]))]);
    summary.push(['Доставка', (oq(26) || 0) + ' / ' + fmtRub_(num_(row[27]))]);
  }
  var cour = ss.getSheetByName(SHEET_COURIERS).getDataRange().getDisplayValues().filter(function (r) { return r[0] === id; })
    .map(function (r) { return { name: r[3], from: r[4], to: r[5], hours: r[6], qty: r[7], orders: r[8], total: r[9], w: r[10], other: r[11], zones: r[12] || '', exSum: r[13] || '',
      extras: extraRows_().filter(function (x) { return String(x[0]) === id && String(x[4]) === String(r[3]); })
        .map(function (x) { return { kind: String(x[5]), zone: String(x[6]), reason: String(x[7]), n: Number(x[8]) || 1, sum: Number(x[9]) || 0 }; }) }; });
  var money = ss.getSheetByName(SHEET_MONEY).getDataRange().getDisplayValues().filter(function (r) { return r[0] === id; })
    .map(function (r) { return { type: r[3], date: r[4], art: r[5], sum: r[6], w: r[7] }; });
  var f = fixFind_(id);
  var rev = {};
  ['r_cash', 'r_card', 'r_ye', 'r_mp', 'r_chib', 'r_sber', 'r_bon'].forEach(function (k, n) { rev[k] = row[5 + n] === 0 ? '' : String(row[5 + n]); });
  var dk = dkey_(row[1]);
  var raw = { date: dk ? String(dk).replace(/(\d{4})(\d{2})(\d{2})/, '$1-$2-$3') : '', branch: String(row[2]), r041: String(row[12] === '' ? '' : row[12]),
    bfact: row[23] === '' || row[23] === undefined || row[23] === null ? '' : String(row[23]), rev: rev,
    ord: { q_pn: oq(24), q_ps: oq(25), q_dn: oq(26), q_ds: oq(27) } };
  return { id: id, own: String(row[3]).trim() === admin.name, admin: String(row[3]), replaced: String(row[22] || ''),
    fix: f ? { status: f.status, rid: f.rid, reason: f.reason } : null, raw: raw, summary: summary, cour: cour, money: money };
}

// ---------- Исправление отправленной смены — только с разрешения руководителя
function fixSheet_() {
  var ss = ss_();
  var sh = ss.getSheetByName(SHEET_FIX);
  if (!sh) {
    sh = ss.insertSheet(SHEET_FIX);
    sh.appendRow(HEADERS[SHEET_FIX]);
    sh.getRange(1, 1, 1, HEADERS[SHEET_FIX].length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  var rep = ss.getSheetByName(SHEET_REPORTS);
  if (!String(rep.getRange('W1').getValue()).trim()) rep.getRange('W1').setValue('Заменён отчётом').setFontWeight('bold');
  return sh;
}
function fixRows_() {
  var sh = ss_().getSheetByName(SHEET_FIX);
  return sh ? sh.getDataRange().getValues() : [];
}
function fixObj_(r, i) {
  return { row: i + 1, sent: d_(r[0]), id: String(r[1]), date: d_(r[2]), branch: String(r[3]), admin: String(r[4]),
    reason: String(r[5]), status: String(r[6]), decided: String(r[7]), newId: String(r[8]), rid: String(r[9]) };
}
function fixFind_(id) {
  var rows = fixRows_();
  for (var i = rows.length - 1; i >= 1; i--) if (String(rows[i][1]) === id) return fixObj_(rows[i], i);
  return null;
}
function fixMap_() {
  var rows = fixRows_(), m = {};
  for (var i = 1; i < rows.length; i++) m[String(rows[i][1])] = fixObj_(rows[i], i);
  return m;
}
function fixList_() {
  var rows = fixRows_(), out = [];
  for (var i = rows.length - 1; i >= 1; i--) if (String(rows[i][6]) === 'ожидает') out.push(fixObj_(rows[i], i));
  return out;
}
function reportRow_(id) {
  var rows = ss_().getSheetByName(SHEET_REPORTS).getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) if (String(rows[i][21]) === id) return { idx: i + 1, r: rows[i] };
  return null;
}
function fixRequest_(id, reason, admin) {
  var rr = reportRow_(id);
  if (!rr || String(rr.r[3]).trim() !== admin.name) return { ok: false, error: 'not_found' };
  if (String(rr.r[22] || '')) return { ok: false, error: 'replaced' };
  var cur = fixFind_(id);
  if (cur && (cur.status === 'ожидает' || cur.status === 'разрешено')) return { ok: true, fix: cur };
  var sh = fixSheet_();
  var rid = 'F' + Utilities.formatDate(new Date(), 'Europe/Moscow', 'yyMMddHHmmss') + Math.floor(Math.random() * 90 + 10);
  sh.appendRow([new Date(), id, rr.r[1], rr.r[2], admin.name, reason.slice(0, 500), 'ожидает', '', '', rid]);
  log_('Запрос на исправление ' + id, admin.name, String(rr.r[2]));
  return { ok: true, fix: fixFind_(id) };
}
function fixDecide_(rid, approve, admin) {
  var sh = fixSheet_(), rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][9]) !== rid) continue;
    if (String(rows[i][6]) !== 'ожидает') return { ok: false, error: 'already', status: String(rows[i][6]) };
    var st = approve ? 'разрешено' : 'отклонено';
    sh.getRange(i + 1, 7, 1, 2).setValues([[st, d_(new Date()) + ', ' + admin.name]]);
    log_('Исправление ' + rows[i][1] + ': ' + st, admin.name, String(rows[i][3]));
    return { ok: true, status: st };
  }
  return { ok: false, error: 'not_found' };
}
function fixCheck_(id, admin) {
  var f = fixFind_(id);
  if (!f || f.status !== 'разрешено') return null;
  var rr = reportRow_(id);
  if (!rr || String(rr.r[22] || '')) return null;
  if (!admin.boss && String(rr.r[3]).trim() !== admin.name) return null;
  return { f: f, rr: rr };
}
function fixDone_(fix, newId, admin) {
  var ss = ss_();
  fixSheet_().getRange(fix.f.row, 7).setValue('исправлено');
  fixSheet_().getRange(fix.f.row, 9).setValue(newId);
  var rep = ss.getSheetByName(SHEET_REPORTS);
  rep.getRange(fix.rr.idx, 23).setValue(newId);
  rep.getRange(rep.getLastRow(), 21).setValue('исправление');
  log_('Исправлен отчёт ' + fix.f.id + ' → ' + newId, admin.name, '');
}
// ---------- Черновик администратора — копия в таблице, чтобы не потерять при сбое компьютера
function draftSheet_() {
  var ss = ss_();
  var sh = ss.getSheetByName(SHEET_DRAFTS);
  if (!sh) {
    sh = ss.insertSheet(SHEET_DRAFTS);
    sh.appendRow(HEADERS[SHEET_DRAFTS]);
    sh.getRange(1, 1, 1, 3).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('C:C').setNumberFormat('@');
  }
  return sh;
}
function draftRow_(sh, name) {
  var n = sh.getLastRow();
  if (n < 2) return 0;
  var v = sh.getRange(2, 1, n - 1, 1).getValues();
  for (var i = 0; i < v.length; i++) if (String(v[i][0]).trim() === name) return i + 2;
  return 0;
}
function draftGet_(admin) {
  var sh = ss_().getSheetByName(SHEET_DRAFTS);
  if (!sh) return null;
  var r = draftRow_(sh, admin.name);
  if (!r) return null;
  var t = String(sh.getRange(r, 3).getValue() || '');
  if (!t) return null;
  try { return JSON.parse(t); } catch (e) { return null; }
}
function draftSave_(admin, d) {
  // пустой черновик = отметка «очищен» со временем, чтобы другие устройства не вернули старую копию
  var txt = JSON.stringify(d || { ts: new Date().getTime(), cleared: true });
  if (txt.length > 45000) return { ok: false, error: 'too_big' };
  var sh = d ? draftSheet_() : ss_().getSheetByName(SHEET_DRAFTS);
  if (!sh) return { ok: true };
  var r = draftRow_(sh, admin.name);
  if (!r) { if (d) sh.appendRow([admin.name, new Date(), txt]); }
  else sh.getRange(r, 2, 1, 2).setValues([[new Date(), txt]]);
  return { ok: true };
}

// Уже есть действующий отчёт за эту дату по этому филиалу?
function dupExists_(r, admin) {
  var branch = r.branch || admin.branch, k = dkey_(r.dateRu);
  if (!k) return false;
  var rows = ss_().getSheetByName(SHEET_REPORTS).getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (dkey_(rows[i][1]) === k && String(rows[i][2]) === branch && !String(rows[i][22] || '')) return true;
  }
  return false;
}

// Заказы курьера по зонам доставки — текст с сайта, например «Жёлтая 0–2,5 км × 3; Синяя 5–8 км × 1»
// Изъятие у курьера: да / нет (сейф) / частично: N ₽ с изъятием, M ₽ без
function wText_(c, t) {
  if (c.w === 'p') { var ws = Math.min(num_(c.ws), t); return 'частично: ' + ws + ' ₽ с изъятием, ' + (t - ws) + ' ₽ без'; }
  return c.w ? 'да' : 'нет (сейф)';
}
function zonesText_(c) { return String((c && c.zt) || '').slice(0, 500); }
function ensureZonesHeader_() {
  var cs = ss_().getSheetByName(SHEET_COURIERS);
  if (cs && !String(cs.getRange('M1').getValue()).trim()) cs.getRange('M1').setValue('По зонам').setFontWeight('bold');
  if (cs && !String(cs.getRange('N1').getValue()).trim()) cs.getRange('N1').setValue('Доплаты').setFontWeight('bold');
  if (cs && !String(cs.getRange('O1').getValue()).trim()) cs.getRange('O1').setValue('Доплаты подробно').setFontWeight('bold');
}
// ---------- Доплаты курьерам: довозы (зона + причина) и прочие доплаты
// ---------- Остатки касс: остаток на начало дня по филиалу (вносит руководитель), дальше сайт считает сам
function cashSheet_() {
  var ss = ss_();
  var sh = ss.getSheetByName(SHEET_CASH);
  if (!sh) {
    sh = ss.insertSheet(SHEET_CASH);
    sh.appendRow(['Филиал', 'Дата', 'Остаток на начало дня', 'Кто внёс', 'Когда']);
    sh.getRange(1, 1, 1, 5).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
function cashStarts_() {
  var sh = ss_().getSheetByName(SHEET_CASH), m = {};
  if (!sh || sh.getLastRow() < 2) return [];
  sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues().forEach(function (r) {
    var b = String(r[0]).trim(), k = dkey_(r[1]);
    if (b && k) m[b] = { branch: b, dk: k, amount: num_(r[2]) }; // последняя запись по филиалу главнее
  });
  return Object.keys(m).map(function (b) { return m[b]; });
}
function cashGet_(admin) {
  var starts = cashStarts_();
  var from = starts.reduce(function (a, x) { return Math.min(a, x.dk); }, 99999999);
  var shifts = starts.length ? history_(admin, { from: String(from), to: '99999999', branch: '' }) : [];
  return { ok: true, starts: starts, shifts: shifts };
}
function cashStart_(b, admin) {
  var branch = String(b.branch || '').trim(), k = Number(String(b.date || '').replace(/\D/g, '')) || 0, amt = num_(b.amount);
  if (!branch || k < 20200101) return { ok: false, error: 'bad_input' };
  var p2 = function (n) { return (n < 10 ? '0' : '') + n; };
  var ds = p2(k % 100) + '.' + p2(Math.floor(k / 100) % 100) + '.' + Math.floor(k / 10000);
  var sh = cashSheet_();
  sh.getRange('B:B').setNumberFormat('@');
  sh.appendRow([branch, ds, amt, admin.name, new Date()]);
  log_('Остаток кассы на ' + ds + ': ' + amt, admin.name, branch);
  return { ok: true };
}

function extraSheet_() {
  var ss = ss_();
  var sh = ss.getSheetByName(SHEET_EXTRA);
  if (!sh) {
    sh = ss.insertSheet(SHEET_EXTRA);
    sh.appendRow(HEADERS[SHEET_EXTRA]);
    sh.getRange(1, 1, 1, HEADERS[SHEET_EXTRA].length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
function extraRows_() {
  var sh = ss_().getSheetByName(SHEET_EXTRA);
  return sh ? sh.getDataRange().getValues().slice(1) : [];
}
// по каждому отчёту: доплаты по видам и довозы по причинам — { kinds: {вид: [кол-во, сумма]}, redo: {причина: [кол-во, сумма]} }
function redoMap_() {
  var m = {};
  extraRows_().forEach(function (r) {
    var id = String(r[0]), a = m[id] = m[id] || { kinds: {}, redo: {} };
    var kind = String(r[5]) || 'Другое', n = Number(r[8]) || 1, s = Number(r[9]) || 0;
    a.kinds[kind] = a.kinds[kind] || [0, 0]; a.kinds[kind][0] += n; a.kinds[kind][1] += s;
    if (kind === 'Довоз') { var k = String(r[7]) || '—'; a.redo[k] = a.redo[k] || [0, 0]; a.redo[k][0] += n; a.redo[k][1] += s; }
  });
  return m;
}
// по каждому отчёту: курьеры — [[ФИО, за часы, заказов, за заказы, итого, доплаты], …] — для статистики руководителя
function courMap_() {
  var sh = ss_().getSheetByName(SHEET_COURIERS), m = {};
  if (!sh || sh.getLastRow() < 2) return m;
  sh.getRange(2, 1, sh.getLastRow() - 1, 14).getValues().forEach(function (r) {
    var id = String(r[0]);
    if (!id) return;
    (m[id] = m[id] || []).push([String(r[3]).trim(), num_(r[6]), num_(r[7]), num_(r[8]), num_(r[9]), num_(r[13])]);
  });
  return m;
}
// по каждому отчёту: строки прихода и расхода — [[тип, статья, сумма], …]
function moneyMap_() {
  var m = {}, sh = ss_().getSheetByName(SHEET_MONEY);
  if (!sh) return m;
  sh.getDataRange().getValues().slice(1).forEach(function (r) {
    var id = String(r[0]), s = num_(r[6]);
    if (!id || !s) return;
    (m[id] = m[id] || []).push([String(r[3]), String(r[5] || '').trim(), s]);
  });
  return m;
}

function num_(v) {
  var n = parseFloat(String(v === undefined || v === null ? '' : v).replace(/\s/g, '').replace(',', '.'));
  return isFinite(n) ? n : 0;
}

// Новые названия видов оплаты и столбец «Безнал факт» в листе «Отчёты» (для уже работающей таблицы)
function ensureReportHeaders_(rep) {
  var h = rep.getRange('F1:I1').getValues()[0];
  if (String(h[0]) === 'Наличка') rep.getRange('F1:I1').setValues([['Наличные', 'Карта', 'Яндекс/Деливери', 'Онлайн']]);
  if (!String(rep.getRange('X1').getValue()).trim()) rep.getRange('X1').setValue('Безнал факт (по сверкам)').setFontWeight('bold');
  if (!String(rep.getRange('Y1').getValue()).trim()) rep.getRange('Y1:AB1').setValues([['Самовывоз, заказов', 'Самовывоз, ₽', 'Доставка, заказов', 'Доставка, ₽']]).setFontWeight('bold');
}
// Самовывоз / доставка: количество и сумма заказов (пусто — если не заполнили)
function ordVal_(r, k) {
  var v = r.ord && r.ord[k];
  return v === undefined || v === null || String(v).trim() === '' ? '' : num_(v);
}
function ordText_(o) {
  o = o || {};
  var has = function (k) { return o[k] !== undefined && o[k] !== null && String(o[k]).trim() !== ''; };
  if (!has('q_pn') && !has('q_ps') && !has('q_dn') && !has('q_ds')) return '';
  var pn = num_(o.q_pn), ps = num_(o.q_ps), dn = num_(o.q_dn), ds = num_(o.q_ds);
  return 'Самовывоз: ' + pn + ' / ' + fmtRub_(ps) + '\nДоставка: ' + dn + ' / ' + fmtRub_(ds) + '\nОбщее: ' + (pn + dn) + ' / ' + fmtRub_(ps + ds);
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
    num_(t.pass), repeat ? 'да' : '', id, '', r.bfact ? num_(r.bfact) : '',
    ordVal_(r, 'q_pn'), ordVal_(r, 'q_ps'), ordVal_(r, 'q_dn'), ordVal_(r, 'q_ds')]);
  ensureReportHeaders_(rep);
  var cs = ss.getSheetByName(SHEET_COURIERS);
  if ((r.cour || []).some(function (c) { return c.zt || (c.ex && c.ex.length); })) ensureZonesHeader_();
  var xs = null;
  (r.cour || []).forEach(function (c) {
    var h = num_(c.hours), o = num_(c.orders), ex = 0, exTxt = [];
    (c.ex || []).forEach(function (x) {
      var s = num_(x.sum); if (!s) return; ex += s;
      var k = x.kind || 'Другое', n = num_(x.n) || 1;
      exTxt.push(k + (x.zone ? ' ' + x.zone : '') + (k === 'Довоз' || k === 'Отказ' || k === 'Листовки' ? ' × ' + n : '') + (x.reason ? ' (' + x.reason + ')' : '') + ' = ' + s);
      if (!xs) xs = extraSheet_();
      xs.appendRow([id, r.dateRu, branch, admin.name, c.name, k, x.zone || '', String(x.reason || '').slice(0, 200), n, s]);
    });
    if (!h && !o && !ex && !c.name) return;
    cs.appendRow([id, r.dateRu, branch, c.name, c.from, c.to, h, num_(c.qty), o, h + o + ex, wText_(c, h + o + ex), c.other ? 'да' : '', zonesText_(c), ex || '', exTxt.join('; ')]);
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

// ---------- Резервная копия таблицы: каждую ночь полная копия в папку на Диске, храним год
var BACKUP_FOLDER = 'Кассовые отчёты — резервные копии';
var BACKUP_KEEP_DAYS = 365;

// Запустить ОДИН раз вручную из редактора: включает ежедневную копию (около 3 ночи по Москве) и сразу делает первую.
function setupBackup() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'backupDaily') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('backupDaily').timeBased().everyDays(1).atHour(3).inTimezone('Europe/Moscow').create();
  backupDaily();
}

function backupDaily() {
  var it = DriveApp.getFoldersByName(BACKUP_FOLDER);
  var folder = it.hasNext() ? it.next() : DriveApp.createFolder(BACKUP_FOLDER);
  var day = Utilities.formatDate(new Date(), 'Europe/Moscow', 'yyyy-MM-dd');
  var name = 'Кассовые отчёты — копия ' + day;
  // одна копия в день: если уже есть за сегодня — заменяем
  var same = folder.getFilesByName(name);
  while (same.hasNext()) same.next().setTrashed(true);
  DriveApp.getFileById(ss_().getId()).makeCopy(name, folder);
  // старше года — в корзину (только наши копии в этой папке)
  var limit = new Date().getTime() - BACKUP_KEEP_DAYS * 24 * 3600 * 1000;
  var files = folder.getFiles();
  while (files.hasNext()) {
    var f = files.next();
    if (f.getName().indexOf('Кассовые отчёты — копия ') === 0 && f.getDateCreated().getTime() < limit) f.setTrashed(true);
  }
  log_('Резервная копия ' + day, '', '');
}

// РАЗОВАЯ: очистка тестовых данных (22.10.2026). Сначала делает резервную копию таблицы.
function clearTestData() {
  backupDaily();
  var ss = ss_();
  ['Отчёты', 'Курьеры', 'Приход-расход', 'Доплаты курьерам', 'Исправления', 'Черновики', 'Журнал входов'].forEach(function (n) {
    var sh = ss.getSheetByName(n);
    if (!sh) { Logger.log(n + ': нет листа'); return; }
    var last = sh.getLastRow();
    if (last > 1) sh.deleteRows(2, last - 1);
    Logger.log(n + ': удалено строк ' + Math.max(0, last - 1));
  });
}
