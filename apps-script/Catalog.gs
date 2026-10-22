/**
 * Каталог для підбирача: двічі на добу (раз на 12 годин) бере товари з Prom API, розставляє теги
 * і зберігає на аркуші «Каталог». Підбирач отримує його за адресою …/exec?catalog=1
 *
 * Токен Prom зберігається у властивостях скрипту (PROM_TOKEN), у коді його немає.
 * Ручні правки: стовпці «… (вручну)» і «Приховати» на аркуші «Каталог» не перезаписуються.
 */

const CATALOG_SHEET = 'Каталог';
const SHOP_URL = 'https://geekshopukraine.com.ua/ua/p';
const CAT_HEAD = ['ID', 'Артикул', 'Назва', 'Ціна', 'Наявність', 'Група',
  'Всесвіт (авто)', 'Тип (авто)', 'Вік від (авто)', 'Для кого (авто)', 'Фото', 'Посилання',
  'Всесвіт (вручну)', 'Тип (вручну)', 'Вік від (вручну)', 'Для кого (вручну)', 'Приховати', 'Оновлено'];
const MANUAL_FROM = 12; // індекс першого ручного стовпця (M), 0-based

/** Запустіть один раз вручну: створює оновлення раз на 12 годин каталогу. */
function installTriggers() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'refreshCatalog')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('refreshCatalog').timeBased().everyHours(12).create();
  refreshCatalog();
}

/* ---------- сповіщення ---------- */
const NOTIFY_TO = 'mr.apemi@gmail.com';
const TOKEN_EXPIRES = '2027-10-22';   // «Дійсний до» токена в кабінеті Prom
const MIN_ITEMS = 50;                  // менше товарів = щось пішло не так

function notify_(key, subject, body, everyHours) {
  const props = PropertiesService.getScriptProperties();
  const last = Number(props.getProperty('NOTIFY_' + key) || 0);
  if (everyHours && Date.now() - last < everyHours * 3600e3) return;   // не спамити
  MailApp.sendEmail(NOTIFY_TO, '[Підбирач Geekshop] ' + subject, body +
    '\n\nТаблиця: ' + SpreadsheetApp.getActive().getUrl() +
    '\nСкрипт: Розширення → Apps Script → Виконання (журнал помилок).');
  props.setProperty('NOTIFY_' + key, String(Date.now()));
}

/** Запустіть вручну, щоб перевірити, що листи доходять. */
function testNotify() {
  notify_('TEST', 'Перевірка сповіщень', 'Це тестовий лист. Якщо він прийшов, сповіщення про помилки працюють.', 0);
}

function refreshCatalog() {
  const props = PropertiesService.getScriptProperties();
  try {
    const count = refreshCatalogCore_();
    if (props.getProperty('CATALOG_FAILING') === '1') {
      notify_('RECOVERED', 'Каталог знову оновлюється', 'Оновлення каталогу з Prom знову працює. Товарів у підбирачі: ' + count + '.', 0);
    }
    props.setProperty('CATALOG_FAILING', '0');
  } catch (err) {
    props.setProperty('CATALOG_FAILING', '1');
    let hint = '';
    const msg = String(err && err.message || err);
    if (/Prom API 401|Prom API 403/.test(msg)) hint = '\n\nСхоже, токен Prom недійсний або без прав «Продукти: Читання». Створіть новий токен у кабінеті Prom і вставте його у властивість PROM_TOKEN (Apps Script → Налаштування проекту).';
    else if (/PROM_TOKEN/.test(msg)) hint = '\n\nУ властивостях скрипту немає PROM_TOKEN.';
    else if (/Prom API 5|Timeout|timed out|Address unavailable/i.test(msg)) hint = '\n\nСхоже на тимчасовий збій Prom. Якщо наступне оновлення (через 12 годин) пройде, прийде лист «Каталог знову оновлюється».';
    notify_('FAIL', 'Помилка оновлення каталогу', 'Не вдалося оновити каталог з Prom.\nПомилка: ' + msg + hint +
      '\n\nПідбирач і далі працює зі старим каталогом.', 12);
    console.error(err);
    return;
  }
  // нагадування про токен за 14 днів до кінця дії
  const daysLeft = Math.floor((new Date(TOKEN_EXPIRES) - new Date()) / 864e5);
  if (daysLeft <= 14) {
    notify_('TOKEN', 'Токен Prom скоро закінчиться', 'Токен «Хелпер подарунків» діє до ' + TOKEN_EXPIRES + ' (залишилось днів: ' + daysLeft + ').\nСтворіть новий токен у кабінеті Prom і замініть PROM_TOKEN у налаштуваннях Apps Script, а потім змініть дату TOKEN_EXPIRES у коді.', 72);
  }
}

function refreshCatalogCore_() {
  const token = PropertiesService.getScriptProperties().getProperty('PROM_TOKEN');
  if (!token) throw new Error('Немає PROM_TOKEN у властивостях скрипту');

  const products = [];
  let lastId = null;
  for (let page = 0; page < 50; page++) {
    let url = 'https://my.prom.ua/api/v1/products/list?limit=100';
    if (lastId) url += '&last_id=' + lastId;
    const res = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + token, 'X-LANGUAGE': 'uk' },
      muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 200) throw new Error('Prom API ' + res.getResponseCode() + ': ' + res.getContentText().slice(0, 300));
    const batch = JSON.parse(res.getContentText()).products || [];
    products.push.apply(products, batch);
    if (batch.length < 100) break;
    lastId = Math.min.apply(null, batch.map(p => p.id)) - 1;
  }

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CATALOG_SHEET) || ss.insertSheet(CATALOG_SHEET);
  // зберегти ручні правки
  const manual = {};
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, CAT_HEAD.length).getValues().forEach(r => {
      manual[String(r[0])] = r.slice(MANUAL_FROM, MANUAL_FROM + 5);
    });
  }

  const now = Utilities.formatDate(new Date(), 'Europe/Kyiv', 'dd.MM.yy HH:mm');
  const seen = {};
  const rows = [];
  products.forEach(p => {
    if (p.status && p.status !== 'on_display') return;
    if (p.presence !== 'available' && p.presence !== 'order') return;
    const name = (p.name_multilang && p.name_multilang.uk) || p.name || '';
    const price = Number(p.price) || 0;
    if (!name || price <= 0) return;
    const key = name.trim().toLowerCase() + '|' + price;
    if (seen[key]) return;
    seen[key] = 1;
    const o = {
      name: name, price: price,
      group: (p.group && (p.group.name_multilang && p.group.name_multilang.uk || p.group.name)) || '',
      desc: String(p.description || '').replace(/<[^>]+>/g, ' ').slice(0, 400)
    };
    o.universe = universe_(o);
    o.type = fixType_(o, type_(o));
    o.age = minAge_(o);
    o.gender = gender_(o);
    const img = thumb_(p.main_image || '');
    const m = manual[String(p.id)] || ['', '', '', '', ''];
    rows.push([String(p.id), p.sku || '', name, price, p.presence, o.group,
      o.universe, o.type, o.age, o.gender, img, SHOP_URL + p.id + '-gift.html'].concat(m, [now]));
  });

  const prev = Number(PropertiesService.getScriptProperties().getProperty('CATALOG_COUNT') || 0);
  if (rows.length < MIN_ITEMS || (prev && rows.length < prev * 0.5)) {
    throw new Error('Prom повернув підозріло мало товарів у наявності: ' + rows.length + ' (було ' + prev + '). Старий каталог залишено без змін.');
  }

  sh.clearContents();
  sh.getRange(1, 1, 1, CAT_HEAD.length).setValues([CAT_HEAD]).setFontWeight('bold');
  sh.setFrozenRows(1);
  if (rows.length) sh.getRange(2, 1, rows.length, CAT_HEAD.length).setValues(rows);
  PropertiesService.getScriptProperties().setProperty('CATALOG_UPDATED', now);
  PropertiesService.getScriptProperties().setProperty('CATALOG_COUNT', String(rows.length));
  return rows.length;
}

/** JSON для підбирача */
function catalogResponse_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CATALOG_SHEET);
  const items = [];
  if (sh && sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, CAT_HEAD.length).getValues().forEach(r => {
      if (r[16] === true || String(r[16]).toLowerCase() === 'так' || String(r[16]) === '1') return;
      const pick = (manualV, autoV) => (manualV === '' || manualV === null) ? autoV : manualV;
      items.push({
        id: String(r[0]), code: String(r[1]), name: String(r[2]), price: Number(r[3]),
        universe: String(pick(r[12], r[6])), type: String(pick(r[13], r[7])),
        age: Number(pick(r[14], r[8])) || 0, gender: String(pick(r[15], r[9])),
        thumb: String(r[10]), url: String(r[11])
      });
    });
  }
  const body = JSON.stringify({
    updated: PropertiesService.getScriptProperties().getProperty('CATALOG_UPDATED') || '',
    items: items
  });
  return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- автотеги (ті самі правила, що в прототипі) ---------- */
function n_(s) { return String(s || '').toLowerCase().replace(/ґ/g, 'г'); }

function universe_(o) {
  const t = n_(o.name + ' ' + o.group);
  const rules = [
    ['stalker', /сталкер|stalker/], ['squid', /кальмар|дальгона/], ['rick', /рік та морті|рік і морті|rick/],
    ['aot', /титан|titan/], ['jjk', /магічна битва|jujutsu/], ['got', /престол/], ['homealone', /один вдома|кевін|маккаліст|filthy/],
    ['fortnite', /фортнайт|fortnite/], ['starwars', /зоряні війни|star wars/],
    ['dc', /бетм|batman|супермен|superman|харлі|harley|харли/],
    ['marvel', /марвел|marvel|залізна людина|iron man|локі|людина-павук|людина павук|халк|грут|дедпул|deadpool|росомаха|зимов|капітан америка|месники|веном/],
    ['hp', /гаррі|поттер|harry|potter|грифіндор|гриффіндор|слизерин|хаффлпафф|пуффендуй|пуфендуй|гафелпаф|рейвенкл|когтевран|хогвартс|гогвортс|дамблдор|герміон|снітч|снитч|дари смерті|мародер|патронус|драко|мандрагор|берті|felix|маховик|пожирач|дементор|доббі|хагрід|фантастичні тварини|мітка|мантія|олліванд|ollivand|заклинань|еліксир|сова|мітла|змія|codex|tria noctis|книга|палич|donum|факультет|gryffindor|полумна|невілл|сортувальна|magic/],
    ['parents', /мама|тато|папа/]
  ];
  for (const [u, re] of rules) if (re.test(t)) return u;
  return 'other';
}

function type_(o) {
  const t = n_(o.name);
  const rules = [
    ['pack', /паковання|пакування|подарункова коробка|бокс•подарункова/], ['costume', /мантія|сортувальна шляпа/],
    ['box', /бокс|набір|набор/], ['wand', /паличк/], ['candle', /свічк/], ['jewelry', /кулон|підвіск/],
    ['keychain', /брелок/], ['apparel', /шапка|рукавич|шарф|краватк/], ['socks', /шкарпет|носоч/],
    ['mug', /чашка|кружка/], ['sweets', /шоколад|цукерк|соти|печиво|жабка|солодощ/], ['drink', /еліксир|напій|напою/],
    ['figure', /фігурк|іграшк/], ['poster', /постер/], ['musicbox', /скриньк/], ['card', /листівк|id картка/],
    ['stickers', /стікер|стикер|наліпк|тату/], ['stationery', /блокнот|ручка|закладк/],
    ['props', /лист|карта|квиток|посвідчення|патронус|окуляри/]
  ];
  for (const [ty, re] of rules) if (re.test(t)) return ty;
  return 'other';
}

function fixType_(o, ty) {
  const t = n_(o.name);
  if (/кулон|підвіск/.test(t)) return 'jewelry';
  if (/мантія|мантією|сортувальна/.test(t)) return 'costume';
  if (/набір слизерин та шарф/.test(t)) return 'costume';
  if (t.indexOf('підставка') >= 0) return 'props';
  if (ty === 'box' && o.price <= 260 && /бокс/.test(t) && t.indexOf('шкарпет') < 0) return 'pack';
  if (ty === 'box') {
    if (/палич/.test(t)) return 'wand';
    if (/шкарпет/.test(t)) return 'socks';
    if (/свічок/.test(t)) return 'candle';
    if (/листів/.test(t)) return 'card';
    if (/краваток/.test(t)) return 'apparel';
    if (/еліксир/.test(t)) return 'drink';
    if (/сот|шоколадок/.test(t)) return 'sweets';
    if (/(набір|набор) (грифіндор|слизерин|хаффлпафф|рейвенклов)/.test(t)) return 'apparel';
  }
  if (ty === 'pack' && o.price > 300) return 'box';
  return ty;
}

function minAge_(o) {
  const t = n_(o.name);
  const base = { candle: 14, jewelry: 8, keychain: 6, drink: 6, sweets: 3, wand: 5, figure: 4, poster: 6, musicbox: 3, mug: 4,
    stationery: 5, stickers: 4, props: 6, card: 0, costume: 5, apparel: 3, socks: 3, pack: 0, box: 6 };
  let a = base.hasOwnProperty(o.type) ? base[o.type] : 6;
  if (/тату|мітка|пожирач/.test(t)) a = Math.max(a, 12);
  if (/боби/.test(t)) a = Math.max(a, 6);
  if (['squid', 'rick', 'stalker'].indexOf(o.universe) >= 0) a = Math.max(a, 16);
  if (['aot', 'jjk', 'got'].indexOf(o.universe) >= 0 || /дедпул|deadpool|харлі/.test(t)) a = Math.max(a, 14);
  if (/водка/.test(t)) a = 18;
  if (o.type === 'box' && /свіч/.test(n_(o.desc))) a = Math.max(a, 14);
  return a;
}

function gender_(o) {
  const nm = n_(o.name);
  if (/для неї/.test(nm)) return 'f';
  if (/для його|для нього/.test(nm)) return 'm';
  const t = n_(o.name + ' ' + o.desc.slice(0, 200));
  if (/для неї|амортенц|кохання|charm of love|герміон/.test(t)) return 'f';
  if (/для його|тато|папа/.test(t)) return 'm';
  if (/мама/.test(t)) return 'f';
  return 'any';
}

function thumb_(u) {
  const m = String(u).match(/^(https:\/\/images\.prom\.ua\/\d+)_(?:w\d+_h\d+_)?(.*)$/);
  return m ? m[1] + '_w320_h320_' + m[2] : u;
}
