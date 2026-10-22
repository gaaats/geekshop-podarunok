/**
 * Geekshop — приймач заявок підбирача подарунків.
 * Вставте цей код у таблицю «Заявки підбирача Geekshop»: Розширення → Apps Script.
 * Потім: Розгорнути → Нове розгортання → Вебзастосунок,
 *   «Виконувати від імені»: Я; «Хто має доступ»: Усі.
 * Скопійоване посилання вебзастосунку (…/exec) надішліть Claude.
 */

const SHEET_NAME = 'Заявки';
const INSTAGRAM = 'geekshop_ukraine';
const SOURCE = 'Підбирач подарунків';

function doGet(e) {
  if (e && e.parameter && e.parameter.catalog) return catalogResponse_();
  let order;
  try {
    const raw = (e && e.parameter && e.parameter.d) || '';
    const json = Utilities.newBlob(Utilities.base64Decode(raw)).getDataAsString('UTF-8');
    order = JSON.parse(json);
  } catch (err) {
    return page_('Не вдалося прочитати заявку', 'Поверніться на сторінку підбору і спробуйте ще раз або напишіть нам в Instagram.', null);
  }

  const clean = sanitize_(order);
  if (!clean) {
    return page_('Заявка неповна', 'Не вистачає імені, телефону, відділення або товарів. Поверніться на сторінку підбору і заповніть форму.', null);
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
    // повторне відкриття того самого посилання не дублює заявку
    const last = sh.getLastRow();
    if (last > 1) {
      const ids = sh.getRange(2, 1, last - 1, 1).getDisplayValues().flat();
      if (ids.indexOf(clean.n) === -1) append_(sh, clean);
    } else {
      append_(sh, clean);
    }
  } finally {
    lock.releaseLock();
  }
  return page_('Заявку ' + clean.n + ' записано', null, clean);
}

function sanitize_(o) {
  if (!o || typeof o !== 'object') return null;
  const s = (v, max) => String(v == null ? '' : v).replace(/[\r\n\t]+/g, ' ').trim().slice(0, max);
  const n = s(o.n, 30);
  const name = s(o.name, 120), phone = s(o.phone, 30).replace(/[^\d+]/g, ''), np = s(o.np, 300);
  if (!/^GS-[\w-]{4,24}$/.test(n) || !name || phone.replace(/\D/g, '').length < 10 || !np) return null;
  const items = (Array.isArray(o.items) ? o.items : []).slice(0, 25).map(it => ({
    c: s(it.c, 40), nm: s(it.nm, 200), p: Math.max(0, Number(it.p) || 0), q: Math.max(1, Math.round(Number(it.q) || 1))
  })).filter(it => it.nm);
  if (!items.length) return null;
  const total = items.reduce((a, it) => a + it.p * it.q, 0);
  return { n, name, phone, np, note: s(o.note, 500), req: s(o.req, 500), items, total };
}

function append_(sh, o) {
  const date = Utilities.formatDate(new Date(), 'Europe/Kyiv', 'dd.MM.yy HH:mm');
  const phone = o.phone.replace(/^\+/, '');
  const money = v => v.toFixed(2);
  const rows = o.items.map(it => [
    o.n, date, o.name, '', phone, '', o.np, 'Нова Пошта', '', 'Нове', '', '',
    it.c, it.nm, money(it.q), SOURCE, 'підбирач', '', '', '',
    money(o.total), money(it.p), money(o.total),
    o.note, o.req, ''
  ]);
  const start = sh.getLastRow() + 1;
  const range = sh.getRange(start, 1, rows.length, rows[0].length);
  range.setNumberFormat('@');
  range.setValues(rows);
}

function page_(title, message, o) {
  const esc = v => String(v).replace(/[&<>"]/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[ch]));
  let text = '';
  if (o) {
    text = ['Заявка ' + o.n, 'ОДЕРЖУВАЧ', "Ім'я: " + o.name, 'Телефон: ' + o.phone, 'Нова пошта: ' + o.np]
      .concat(o.note ? ['Побажання: ' + o.note] : [])
      .concat(['', 'СКЛАД НАБОРУ'])
      .concat(o.items.map(it => '• ' + it.nm + (it.q > 1 ? ' ×' + it.q : '') + ' — ' + it.p + ' ₴'))
      .concat(['', 'Разом: ' + o.total + ' ₴ (без доставки)']).join('\n');
  }
  const html = `<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<style>
body{margin:0;background:#0D0A16;color:#F2EEFF;font:16px/1.55 system-ui,-apple-system,"Segoe UI",sans-serif;padding:28px 18px}
.w{max-width:560px;margin:0 auto;display:grid;gap:18px}
h1{font-size:26px;line-height:1.2;margin:0}
.st{display:grid;gap:12px;counter-reset:s}
.s{display:grid;grid-template-columns:34px 1fr;gap:12px;align-items:start;background:#16112A;border:1px solid #2F2750;border-radius:14px;padding:14px}
.s b{width:34px;height:34px;border-radius:50%;background:#F0C25A;color:#0D0A16;display:grid;place-items:center;font-weight:800}
.btn{display:block;text-align:center;font-weight:700;text-decoration:none;border-radius:14px;padding:15px;background:#EC3D52;color:#fff;border:none;font-size:16px;width:100%;cursor:pointer}
.btn.g{background:#F0C25A;color:#0D0A16}
pre{white-space:pre-wrap;background:#16112A;border:1px dashed #2F2750;border-radius:12px;padding:14px;font:14px/1.5 ui-monospace,Menlo,monospace;margin:0}
p{margin:0;color:#A79DC9}
</style></head><body><div class="w">
<h1>${esc(title)}</h1>
${message ? `<p>${esc(message)}</p><a class="btn" href="https://ig.me/m/${INSTAGRAM}" target="_blank">Написати нам в Instagram</a>` : `
<p>Залишився один крок: надішліть цей текст нам у Direct, і менеджер підтвердить замовлення та надішле реквізити для оплати.</p>
<div class="st">
<div class="s"><b>1</b><div>Скопіюйте текст заявки.<br><button class="btn g" id="cp" style="margin-top:10px">Скопіювати заявку</button></div></div>
<div class="s"><b>2</b><div>Відкрийте наш Instagram <strong>@${INSTAGRAM}</strong>.<br><a class="btn" style="margin-top:10px" href="https://ig.me/m/${INSTAGRAM}" target="_blank">Відкрити Direct</a></div></div>
<div class="s"><b>3</b><div>Вставте текст у повідомлення (утримайте поле введення → «Вставити») і надішліть.</div></div>
</div>
<pre id="t">${esc(text)}</pre>`}
</div>
<script>
var b=document.getElementById('cp');
if(b) b.onclick=function(){var t=document.getElementById('t').innerText;
  (navigator.clipboard?navigator.clipboard.writeText(t):Promise.reject()).then(function(){b.textContent='✓ Скопійовано';})
  .catch(function(){var r=document.createRange();r.selectNodeContents(document.getElementById('t'));var s=getSelection();s.removeAllRanges();s.addRange(r);b.textContent='Текст виділено: скопіюйте його';});};
</script></body></html>`;
  return HtmlService.createHtmlOutput(html).setTitle('Заявка Geekshop')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
