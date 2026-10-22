// Посредник «Google-таблица → MAX». Пересылает запрос ТОЛЬКО на platform-api2.max.ru
// и только такие запросы: проверка бота, сведения о чате, последние события бота (чтобы узнать ID чата), отправка сообщения в чат.
// Ключ бота сюда не записывается: его присылает скрипт таблицы в заголовке X-Max-Auth.
const https = require('https');
const tls = require('tls');

// У MAX сертификат российского УЦ Минцифры. Его корневой и промежуточный сертификаты
// скачиваем с сайта Госуслуг (gu-st.ru) при запуске и добавляем к обычным — проверка сертификата остаётся включённой.
const CA_URLS = ['https://gu-st.ru/content/lending/russian_trusted_root_ca_pem.crt',
                 'https://gu-st.ru/content/lending/russian_trusted_sub_ca_pem.crt'];
let CA = null;
function get(url) {
  return new Promise((resolve, reject) => {
    https.get(url, (x) => { let s = ''; x.setEncoding('utf8'); x.on('data', (c) => { s += c; }); x.on('end', () => resolve(s)); })
      .on('error', reject).setTimeout(10000, function () { this.destroy(new Error('timeout')); });
  });
}
async function caList() {
  if (CA) return CA;
  const extra = await Promise.all(CA_URLS.map(get));
  if (!extra.every((p) => p.includes('BEGIN CERTIFICATE'))) throw new Error('не удалось скачать сертификаты Минцифры');
  CA = tls.rootCertificates.concat(extra);
  return CA;
}

module.exports.handler = async (event) => {
  const h = event.headers || {};
  const token = h['X-Max-Auth'] || h['x-max-auth'];
  if (!token) return { statusCode: 401, body: 'no token' };

  let q = {};
  try {
    q = JSON.parse(event.isBase64Encoded ? Buffer.from(event.body || '', 'base64').toString('utf8') : (event.body || '{}'));
  } catch (e) { return { statusCode: 400, body: 'bad json' }; }

  const method = String(q.m || 'GET').toUpperCase();
  const path = String(q.p || '');
  const allowed = (method === 'GET' && /^\/(me|chats(\?count=\d+)?|chats\/-?\d+|updates(\?[a-z_=&,\d]*)?)$/.test(path)) ||
                  (method === 'POST' && /^\/messages\?chat_id=-?\d+$/.test(path));
  if (!allowed) return { statusCode: 400, body: 'not allowed' };

  const data = q.b ? JSON.stringify(q.b) : null;
  try {
    const ca = await caList();
    const res = await new Promise((resolve, reject) => {
      const headers = { Authorization: token };
      if (data) { headers['Content-Type'] = 'application/json'; headers['Content-Length'] = Buffer.byteLength(data); }
      const r = https.request({ host: 'platform-api2.max.ru', path, method, headers, ca }, (x) => {
        let s = ''; x.setEncoding('utf8');
        x.on('data', (c) => { s += c; });
        x.on('end', () => resolve({ code: x.statusCode, body: s }));
      });
      r.on('error', reject);
      r.setTimeout(15000, () => r.destroy(new Error('timeout')));
      if (data) r.write(data);
      r.end();
    });
    return { statusCode: 200, headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(res) };
  } catch (e) {
    return { statusCode: 502, body: 'relay error: ' + e.message };
  }
};
