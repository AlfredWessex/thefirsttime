/**
 * 金煎堂 線上訂單系統（Google Apps Script）
 *
 * 功能：
 *   1. doPost：接收網站送出的訂單 → 驗證、用「商品價目」重新計算金額 → 寫入試算表 → 寄確認信
 *   2. doGet ：用「訂單編號＋手機」查詢訂單狀態（也可直接用瀏覽器打開網址測試是否部署成功）
 *   3. 狀態改成「可取餐」或「已出貨」時，自動寄 Email 通知顧客
 *   4. setup：一鍵建立所有工作表、欄位、下拉選單、顏色標示、統計報表
 *
 * 使用方式請看同資料夾的 README.md。
 */

const TZ = 'Asia/Taipei';
const SHEET = { ORDERS: '訂單', ITEMS: '訂單明細', PRICES: '商品價目', STATS: '每日統計', SETTINGS: '設定' };
const STATUS = ['新訂單', '製作中', '可取餐', '已出貨', '已完成', '已取消'];
const NOTIFY_STATUS = ['可取餐', '已出貨'];
const TYPE_LABEL = { dinein: '內用', takeout: '外帶', delivery: '宅配' };
const PAY_OPTIONS = ['現場付款', 'LINE Pay', '信用卡'];
const STORES = { flagship: '永和旗艦店', xinyi: '信義店', daan: '大安店', taichung: '台中勤美店' };
const SHIPPING_FEE = 160, FREE_SHIPPING = 1500;

const ORDER_HEADERS = ['訂單編號', '建立時間', '狀態', '取餐方式', '門市', '到店時間', '姓名', '手機', 'Email',
  '宅配地址', '付款方式', '商品明細', '商品小計', '運費', '應付金額', '備註', '系統備註'];
const ITEM_HEADERS = ['訂單編號', '建立時間', '商品編號', '品名', '單價', '數量', '小計'];
const PRICE_HEADERS = ['商品編號', '品名', '單價', '類別', '上架中'];
const COL = Object.fromEntries(ORDER_HEADERS.map((h, i) => [h, i + 1]));

// 與網站 index.html 的 PRODUCTS / GIFTS 相同；setup 時寫入「商品價目」工作表
const DEFAULT_PRICES = [
  ['p1', '招牌豬肉鍋貼', 80, '招牌鍋貼'], ['p2', '韭黃鮮蝦鍋貼', 120, '招牌鍋貼'],
  ['p3', '麻辣牛肉鍋貼', 130, '招牌鍋貼'], ['p4', '起司玉米鍋貼', 100, '招牌鍋貼'],
  ['d1', '高麗菜豬肉水餃', 70, '精選水餃'], ['d2', '玉米豬肉水餃', 75, '精選水餃'],
  ['d3', '紅油抄手', 90, '精選水餃'], ['d4', '香菇素食水餃', 75, '精選水餃'],
  ['b1', '酸辣湯', 45, '湯品/飲料'], ['b2', '玉米濃湯', 45, '湯品/飲料'],
  ['b3', '古早味豆漿', 30, '湯品/飲料'], ['b4', '冬瓜檸檬', 40, '湯品/飲料'],
  ['s1', '涼拌小黃瓜', 35, '經典小菜'], ['s2', '滷味拼盤', 60, '經典小菜'],
  ['s3', '麻辣鴨血', 50, '經典小菜'], ['s4', '皮蛋豆腐', 45, '經典小菜'],
  ['g1', '招牌招財鍋貼禮盒', 680, '冷凍禮盒'], ['g2', '闔家團圓綜合禮盒', 980, '冷凍禮盒'],
  ['g3', '節慶福氣伴手禮盒', 880, '冷凍禮盒']
];

/* =========================================================
 * 1. 一鍵初始化（在編輯器選擇 setup 後按「執行」，只需做一次）
 * ======================================================= */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(TZ);

  // 訂單
  const orders = getOrCreateSheet_(ss, SHEET.ORDERS, ORDER_HEADERS);
  orders.getRange('A:A').setNumberFormat('@');
  orders.getRange('B:B').setNumberFormat('yyyy/mm/dd hh:mm');
  orders.getRange('H:H').setNumberFormat('@'); // 手機保留開頭的 0
  orders.getRange('M:O').setNumberFormat('$#,##0');
  [130, 130, 80, 70, 100, 110, 80, 100, 180, 220, 80, 260, 80, 60, 90, 180, 160]
    .forEach((w, i) => orders.setColumnWidth(i + 1, w));
  orders.getRange('L:L').setWrap(true);
  const statusRange = orders.getRange(2, COL['狀態'], orders.getMaxRows() - 1, 1);
  statusRange.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(STATUS, true).setAllowInvalid(false).build());
  const colors = { '新訂單': '#fde68a', '製作中': '#fed7aa', '可取餐': '#bbf7d0', '已出貨': '#bfdbfe', '已完成': '#e5e7eb', '已取消': '#fecaca' };
  orders.setConditionalFormatRules(Object.keys(colors).map(s =>
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(s).setBackground(colors[s]).setRanges([statusRange]).build()));

  // 訂單明細
  const items = getOrCreateSheet_(ss, SHEET.ITEMS, ITEM_HEADERS);
  items.getRange('A:A').setNumberFormat('@');
  items.getRange('B:B').setNumberFormat('yyyy/mm/dd hh:mm');
  items.getRange('E:E').setNumberFormat('$#,##0');
  items.getRange('G:G').setNumberFormat('$#,##0');

  // 商品價目（已有資料就不覆蓋，避免蓋掉你改過的價格）
  const prices = getOrCreateSheet_(ss, SHEET.PRICES, PRICE_HEADERS);
  if (prices.getLastRow() < 2) {
    prices.getRange(2, 1, DEFAULT_PRICES.length, 5).setValues(DEFAULT_PRICES.map(r => r.concat([true])));
  }
  prices.getRange(2, 5, prices.getMaxRows() - 1, 1).insertCheckboxes();
  prices.getRange('C:C').setNumberFormat('$#,##0');

  // 設定
  const settings = getOrCreateSheet_(ss, SHEET.SETTINGS, ['設定項目', '值', '說明']);
  if (settings.getLastRow() < 2) {
    settings.getRange(2, 1, 3, 3).setValues([
      ['店家通知 Email', Session.getEffectiveUser().getEmail(), '有新訂單時通知這個信箱（多個用逗號分隔，留空就不通知）'],
      ['寄送顧客確認信', true, '顧客有填 Email 時，送出訂單後寄確認信'],
      ['狀態通知顧客', true, '訂單狀態改為「可取餐」或「已出貨」時寄信通知顧客']
    ]);
    settings.getRange('B3:B4').insertCheckboxes();
    settings.setColumnWidth(1, 140); settings.setColumnWidth(2, 260); settings.setColumnWidth(3, 420);
  }

  // 每日統計（公式自動計算，不含已取消的訂單）
  const stats = getOrCreateSheet_(ss, SHEET.STATS, []);
  stats.clear();
  stats.getRange('A1').setValue('每日營業額').setFontWeight('bold');
  stats.getRange('A2').setFormula(
    `=IFERROR(QUERY(${SHEET.ORDERS}!A2:O, "select toDate(B), count(A), sum(O) where A is not null and C <> '已取消' ` +
    `group by toDate(B) order by toDate(B) desc label toDate(B) '日期', count(A) '訂單數', sum(O) '營業額'", 0), "尚無訂單")`);
  stats.getRange('E1').setValue('熱銷商品排行').setFontWeight('bold');
  stats.getRange('E2').setFormula(
    `=IFERROR(QUERY(${SHEET.ITEMS}!A2:G, "select D, sum(F), sum(G) where A is not null ` +
    `group by D order by sum(F) desc label D '品名', sum(F) '數量', sum(G) '金額'", 0), "尚無訂單")`);
  stats.getRange('A:A').setNumberFormat('yyyy/mm/dd');
  stats.getRange('C:C').setNumberFormat('$#,##0');
  stats.getRange('G:G').setNumberFormat('$#,##0');

  // 刪掉新試算表預設的空白「工作表1」
  ['工作表1', 'Sheet1'].forEach(n => {
    const s = ss.getSheetByName(n);
    if (s && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s);
  });

  // 安裝「狀態變更 → 寄信」觸發條件（重複執行不會重複安裝）
  const exists = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'onStatusEdit');
  if (!exists) ScriptApp.newTrigger('onStatusEdit').forSpreadsheet(ss).onEdit().create();

  ss.setActiveSheet(orders);
  return '設定完成';
}

function getOrCreateSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (headers.length) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#432a18').setFontColor('#fbf6ee');
    sh.setFrozenRows(1);
  }
  return sh;
}

/* =========================================================
 * 2. 接收訂單
 * ======================================================= */
function doPost(e) {
  try {
    const data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    return json_(createOrder_(data));
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: '系統忙碌中，請稍後再試或來電門市訂購。' });
  }
}

function createOrder_(data) {
  // 機器人防護：網站上看不到的欄位被填了就當作垃圾訂單
  if (data.website) return { ok: false, error: '訂單送出失敗' };

  const name = clean_(data.name, 30);
  const phone = String(data.phone || '').replace(/[\s-]/g, '');
  const email = clean_(data.email, 100);
  const type = data.orderType;
  const address = clean_(data.address, 150);
  const note = clean_(data.note, 300);
  const pay = PAY_OPTIONS.includes(data.pay) ? data.pay : '現場付款';

  if (!name) return { ok: false, error: '請輸入姓名' };
  if (!/^09\d{8}$/.test(phone)) return { ok: false, error: '請輸入 09 開頭的 10 碼手機號碼' };
  if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return { ok: false, error: 'Email 格式不正確' };
  if (!TYPE_LABEL[type]) return { ok: false, error: '請選擇取餐方式' };
  if (type === 'delivery' && address.length < 6) return { ok: false, error: '請輸入完整宅配地址' };
  if (type !== 'delivery' && !STORES[data.storeId]) return { ok: false, error: '請選擇取餐門市' };

  // 用試算表上的價格重新計算，避免網頁上的金額被竄改
  const priceMap = getPriceMap_();
  const lines = [];
  for (const it of (Array.isArray(data.items) ? data.items : [])) {
    const p = priceMap[it.id];
    const qty = Math.floor(Number(it.qty));
    if (!p) return { ok: false, error: '購物車內有已下架的商品，請重新整理頁面' };
    if (!(qty >= 1 && qty <= 99)) return { ok: false, error: '商品數量不正確' };
    lines.push({ id: it.id, name: p.name, price: p.price, qty, sum: p.price * qty });
  }
  if (!lines.length) return { ok: false, error: '購物車是空的' };

  const subtotal = lines.reduce((s, l) => s + l.sum, 0);
  const shipping = type === 'delivery' && subtotal < FREE_SHIPPING ? SHIPPING_FEE : 0;
  const total = subtotal + shipping;
  const priceChanged = Number(data.clientTotal) !== total;

  // 防止重複送出：同一支手機 20 秒內只接受一筆
  const cache = CacheService.getScriptCache();
  if (cache.get('dup_' + phone)) return { ok: false, error: '訂單已送出，請勿重複點擊' };

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let orderNo, now;
  try {
    now = new Date();
    orderNo = nextOrderNo_(now);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const detail = lines.map(l => `${l.name} × ${l.qty}`).join('\n');
    ss.getSheetByName(SHEET.ORDERS).appendRow([
      orderNo, now, '新訂單', TYPE_LABEL[type],
      type === 'delivery' ? '' : STORES[data.storeId],
      type === 'delivery' ? '' : clean_(data.time, 30),
      name, "'" + phone, email, type === 'delivery' ? address : '', pay, detail,
      subtotal, shipping, total, note,
      priceChanged ? `網頁金額 ${data.clientTotal}，已依價目表改為 ${total}` : ''
    ]);
    const itemSheet = ss.getSheetByName(SHEET.ITEMS);
    const rows = lines.map(l => [orderNo, now, l.id, l.name, l.price, l.qty, l.sum]);
    itemSheet.getRange(itemSheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  cache.put('dup_' + phone, '1', 20);

  const order = { orderNo, now, type, name, phone, email, address, note, pay, lines, subtotal, shipping, total,
    store: STORES[data.storeId], time: clean_(data.time, 30) };
  sendMails_(order);

  return { ok: true, orderNo, subtotal, shipping, total, priceChanged };
}

// 訂單編號：JJT + 年月日 + 當天流水號，例如 JJT261127-003
function nextOrderNo_(now) {
  const day = Utilities.formatDate(now, TZ, 'yyMMdd');
  const props = PropertiesService.getScriptProperties();
  const seq = Number(props.getProperty('seq_' + day) || 0) + 1;
  props.setProperty('seq_' + day, String(seq));
  return `JJT${day}-${String(seq).padStart(3, '0')}`;
}

function getPriceMap_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.PRICES);
  const map = {};
  if (sh.getLastRow() < 2) return map;
  sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues().forEach(([id, name, price, , active]) => {
    if (id && active === true && Number(price) > 0) map[String(id).trim()] = { name: String(name), price: Number(price) };
  });
  return map;
}

/* =========================================================
 * 3. 查詢訂單：<網址>?orderNo=JJT261127-001&phone=0912345678
 * ======================================================= */
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (!p.orderNo) return json_({ ok: true, service: '金煎堂訂單系統運作中' });
  const phone = String(p.phone || '').replace(/[\s-]/g, '');
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.ORDERS);
  const last = sh.getLastRow();
  if (last >= 2) {
    const rows = sh.getRange(2, 1, last - 1, ORDER_HEADERS.length).getValues();
    for (const r of rows) {
      if (r[COL['訂單編號'] - 1] === p.orderNo && String(r[COL['手機'] - 1]) === phone) {
        return json_({ ok: true, orderNo: p.orderNo, status: r[COL['狀態'] - 1], total: r[COL['應付金額'] - 1] });
      }
    }
  }
  return json_({ ok: false, error: '查無此訂單，請確認訂單編號與手機號碼' });
}

/* =========================================================
 * 4. 狀態改為「可取餐 / 已出貨」時通知顧客（由 setup 安裝的觸發條件呼叫）
 * ======================================================= */
function onStatusEdit(e) {
  const range = e && e.range;
  if (!range) return;
  const sh = range.getSheet();
  if (sh.getName() !== SHEET.ORDERS || range.getColumn() !== COL['狀態'] || range.getNumRows() !== 1 || range.getRow() < 2) return;
  const status = e.value;
  if (!NOTIFY_STATUS.includes(status) || !getSetting_('狀態通知顧客', true)) return;
  const row = sh.getRange(range.getRow(), 1, 1, ORDER_HEADERS.length).getValues()[0];
  const email = row[COL['Email'] - 1];
  if (!email) return;
  const no = row[COL['訂單編號'] - 1], name = row[COL['姓名'] - 1];
  const msg = status === '可取餐'
    ? `您的餐點已經煎好囉！請至 ${row[COL['門市'] - 1]} 取餐，報上訂單編號即可。`
    : '您訂購的冷凍禮盒已經出貨，預計 1–2 天內送達，收到後請立即冷凍保存。';
  MailApp.sendEmail({
    to: email,
    subject: `【金煎堂】訂單 ${no} ${status}`,
    htmlBody: mailLayout_(`${esc_(name)} 您好：`, `<p>${msg}</p><p>訂單編號：<b>${esc_(no)}</b><br>應付金額：<b>$${row[COL['應付金額'] - 1]}</b></p>`)
  });
  sh.getRange(range.getRow(), COL['系統備註']).setValue(
    `${row[COL['系統備註'] - 1] ? row[COL['系統備註'] - 1] + '\n' : ''}${Utilities.formatDate(new Date(), TZ, 'MM/dd HH:mm')} 已寄「${status}」通知`);
}

/* =========================================================
 * Email
 * ======================================================= */
function sendMails_(o) {
  const typeLabel = TYPE_LABEL[o.type];
  const where = o.type === 'delivery' ? `宅配地址：${esc_(o.address)}` : `取餐門市：${esc_(o.store)}<br>預計到店：${esc_(o.time)}`;
  const table = '<table style="border-collapse:collapse;width:100%;font-size:14px">' +
    o.lines.map(l => `<tr><td style="padding:6px 0;border-bottom:1px solid #ecdfcf">${esc_(l.name)} × ${l.qty}</td>` +
      `<td style="padding:6px 0;border-bottom:1px solid #ecdfcf;text-align:right">$${l.sum}</td></tr>`).join('') +
    `<tr><td style="padding:6px 0">運費</td><td style="text-align:right">${o.shipping ? '$' + o.shipping : '免運'}</td></tr>` +
    `<tr><td style="padding:6px 0;font-weight:bold">應付金額</td><td style="text-align:right;font-weight:bold;color:#cc5f17">$${o.total}</td></tr></table>`;
  const info = `<p>訂單編號：<b>${o.orderNo}</b><br>取餐方式：${typeLabel}<br>${where}<br>付款方式：${esc_(o.pay)}</p>`;

  try {
    if (o.email && getSetting_('寄送顧客確認信', true)) {
      MailApp.sendEmail({
        to: o.email, subject: `【金煎堂】訂單成立通知 ${o.orderNo}`, name: '金煎堂',
        htmlBody: mailLayout_(`${esc_(o.name)} 您好，感謝您的訂購！`, info + table +
          '<p style="color:#8a5a36;font-size:13px">如需修改或取消訂單，請直接來電門市。</p>')
      });
    }
  } catch (err) { console.error('顧客確認信寄送失敗', err); }

  try {
    const to = String(getSetting_('店家通知 Email', '') || '').trim();
    if (to) {
      MailApp.sendEmail({
        to, subject: `🔔 新訂單 ${o.orderNo}｜${typeLabel}｜$${o.total}`,
        htmlBody: mailLayout_('有一筆新訂單', info + `<p>顧客：${esc_(o.name)}　${o.phone}<br>Email：${esc_(o.email || '—')}<br>備註：${esc_(o.note || '—')}</p>` + table +
          `<p><a href="${SpreadsheetApp.getActiveSpreadsheet().getUrl()}">打開訂單試算表</a></p>`)
      });
    }
  } catch (err) { console.error('店家通知信寄送失敗', err); }
}

function mailLayout_(title, body) {
  return `<div style="max-width:520px;margin:auto;font-family:sans-serif;color:#2c1b10">` +
    `<div style="background:#432a18;color:#f8b071;padding:16px 20px;font-size:20px;font-weight:bold;border-radius:12px 12px 0 0">金煎堂 JIN JIAN TANG</div>` +
    `<div style="background:#fbf6ee;padding:20px;border-radius:0 0 12px 12px"><h2 style="font-size:18px;margin:0 0 12px">${title}</h2>${body}</div></div>`;
}

/* =========================================================
 * 小工具
 * ======================================================= */
function getSetting_(key, fallback) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.SETTINGS);
  if (!sh || sh.getLastRow() < 2) return fallback;
  const row = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().find(r => r[0] === key);
  return row ? row[1] : fallback;
}

// 清理文字：去頭尾空白、限制長度，並避免以 = + - @ 開頭被試算表當成公式執行
function clean_(v, max) {
  let s = String(v == null ? '' : v).trim().slice(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function esc_(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
