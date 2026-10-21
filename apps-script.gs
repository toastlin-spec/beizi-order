/**
 * 杯子洪了 團購點餐 — Google Apps Script 後端
 *
 * 安裝步驟：
 * 1. 建立一個新的 Google 試算表 → 選單「擴充功能」→「Apps Script」
 * 2. 把預設的 .gs 檔內容全部刪掉，貼上本檔案，按儲存（檔名不用改）
 * 3. 檔案旁的「＋」→「HTML」→ 命名為 Index，貼上 index.html 全部內容並儲存
 * 4. 右上角「部署」→「新增部署作業」→ 類型選「網頁應用程式」
 *      執行身分：我
 *      誰可以存取：所有人
 * 5. 按「部署」並完成 Google 授權，「網頁應用程式網址」（結尾是 /exec）就是點餐網頁
 *    （若網頁要放在其他網站，則把此網址填到 index.html 的 API_URL）
 *
 * 之後如果修改了這份程式碼，要到「部署」→「管理部署作業」→ 編輯 → 版本選「新版本」再部署，網址不變。
 *
 * 資料格式：一筆訂單（一個人的購物車）佔多列，每個品項一列；
 *          第一列的「原始資料」欄存整筆訂單的 JSON，供網頁讀回。
 */

const SHEET_NAME = '訂單';
const ADMIN_KEY = '6789';   // 需與 index.html 的 ADMIN_PIN 相同
const HEADERS = ['訂單ID', '下單時間', '姓名', '電話', '分類', '飲品', '冰熱', '甜度', '加料',
                 '杯數', '自備杯', '單價', '原價小計', '贈送杯數', '優惠折抵', '應付金額',
                 '訂單應付合計', '原始資料(勿改)'];
const RAW_COL = HEADERS.length;

// 打開網址時顯示點餐網頁（專案中需有名為 Index 的 HTML 檔）
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('杯子洪了 團購點餐')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 網頁放在 Apps Script 時，透過 google.script.run 呼叫這裡
function api(body) {
  return JSON.parse(handle_(body).getContent());
}

// 網頁放在其他網站（設定 API_URL）時，透過 POST 呼叫這裡
function doPost(e) {
  try {
    return handle_(JSON.parse(e.postData.contents));
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function handle_(body) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    switch (body.action) {
      case 'add':
        appendOrder_(body.order);
        return json_({ ok: true });
      case 'list':
        requireKey_(body);
        return json_({ ok: true, orders: readAll_() });
      case 'update':
        requireKey_(body);
        updateOrder_(body.order);
        return json_({ ok: true });
      case 'delete':
        requireKey_(body);
        deleteOrder_(body.id);
        return json_({ ok: true });
      case 'clear':
        requireKey_(body);
        clearAll_();
        return json_({ ok: true });
      default:
        return json_({ ok: false, error: '不支援的動作' });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

/* ---------- 內部函式 ---------- */

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function requireKey_(body) {
  if (body.key !== ADMIN_KEY) throw new Error('團主密碼錯誤');
}

function stamp_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmm');
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  // 欄位格式與目前版本不同（舊版工作表）→ 改名保留，另建新工作表
  if (sh) {
    const header = sh.getRange(1, 1, 1, HEADERS.length).getValues()[0].join('|');
    if (header !== HEADERS.join('|')) {
      sh.setName('舊版_' + stamp_());
      sh = null;
    }
  }
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#C8161D').setFontColor('#FFFFFF');
    sh.hideColumns(RAW_COL);
  }
  return sh;
}

const str_ = (v, n) => String(v == null ? '' : v).slice(0, n);
const num_ = v => Math.max(0, Math.round(Number(v) || 0));

function cleanItem_(it) {
  if (!it || typeof it !== 'object' || !str_(it.drink, 30).trim()) throw new Error('品項資料錯誤');
  return {
    drinkId: str_(it.drinkId, 10),
    cat: str_(it.cat, 20),
    drink: str_(it.drink, 30),
    variant: it.variant ? str_(it.variant, 30) : null,
    base: num_(it.base),
    ice: str_(it.ice, 10),
    sugar: str_(it.sugar, 10),
    toppings: Array.isArray(it.toppings) ? it.toppings.slice(0, 10).map(t => str_(t, 10)) : [],
    toppingCost: num_(it.toppingCost),
    qty: Math.min(30, Math.max(1, num_(it.qty))),
    ownCup: !!it.ownCup,
    unit: num_(it.unit),
    gross: num_(it.gross),
    free: num_(it.free),
    discount: num_(it.discount),
    subtotal: num_(it.subtotal),
  };
}

function clean_(o) {
  if (!o || typeof o !== 'object') throw new Error('訂單資料錯誤');
  if (!str_(o.name, 20).trim() || !str_(o.phone, 15).trim()) throw new Error('姓名、電話為必填');
  if (!Array.isArray(o.items) || !o.items.length || o.items.length > 30) throw new Error('購物車品項數量錯誤');
  return {
    id: str_(o.id, 40) || Utilities.getUuid(),
    time: str_(o.time, 40) || new Date().toISOString(),
    name: str_(o.name, 20).trim(),
    phone: str_(o.phone, 15).trim(),
    items: o.items.map(cleanItem_),
    cups: num_(o.cups),
    gross: num_(o.gross),
    promo: str_(o.promo, 10),
    free: num_(o.free),
    discount: num_(o.discount),
    subtotal: num_(o.subtotal),
  };
}

function toRows_(o) {
  return o.items.map((it, j) => [
    o.id, new Date(o.time), o.name, "'" + o.phone, it.cat, it.drink, it.ice, it.sugar,
    it.toppings.length ? it.toppings.join('、') : '無', it.qty, it.ownCup ? '是' : '否',
    it.unit, it.gross, it.free, it.discount, it.subtotal,
    j === 0 ? o.subtotal : '', j === 0 ? JSON.stringify(o) : '',
  ]);
}

function writeRowsAt_(sh, afterRow, rows) {
  sh.insertRowsAfter(afterRow, rows.length);
  sh.getRange(afterRow + 1, 1, rows.length, HEADERS.length).setValues(rows);
}

function appendOrder_(order) {
  const o = clean_(order);
  o.time = new Date().toISOString();   // 以伺服器時間為準
  const sh = sheet_();
  writeRowsAt_(sh, sh.getLastRow(), toRows_(o));
}

function readAll_() {
  const sh = sheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, RAW_COL, last - 1, 1).getValues()
    .filter(r => r[0] !== '')
    .map(r => { try { return JSON.parse(r[0]); } catch (e) { return null; } })
    .filter(Boolean);
}

// 回傳該訂單佔用的所有列號（由小到大）
function findRows_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 1).getValues()
    .map((r, i) => String(r[0]) === String(id) ? i + 2 : 0)
    .filter(Boolean);
}

// 試算表不允許刪光所有未凍結的列，必要時先補一列空白
function ensureSpareRow_(sh, deleting) {
  if (sh.getMaxRows() - 1 - deleting < 1) sh.insertRowsAfter(sh.getMaxRows(), 1);
}

function removeRows_(sh, rows) {
  ensureSpareRow_(sh, rows.length);
  rows.slice().reverse().forEach(r => sh.deleteRow(r));
}

function updateOrder_(order) {
  const o = clean_(order);
  const sh = sheet_();
  const rows = findRows_(sh, o.id);
  if (!rows.length) throw new Error('找不到這筆訂單，可能已被刪除');
  removeRows_(sh, rows);
  writeRowsAt_(sh, rows[0] - 1, toRows_(o));   // 寫回原本的位置
}

function deleteOrder_(id) {
  const sh = sheet_();
  const rows = findRows_(sh, id);
  if (!rows.length) throw new Error('找不到這筆訂單，可能已被刪除');
  removeRows_(sh, rows);
}

function clearAll_() {
  const sh = sheet_();
  const last = sh.getLastRow();
  if (last < 2) return;
  // 清空前自動封存一份，避免誤刪
  sh.copyTo(SpreadsheetApp.getActiveSpreadsheet()).setName('封存_' + stamp_());
  ensureSpareRow_(sh, last - 1);
  sh.deleteRows(2, last - 1);
}
