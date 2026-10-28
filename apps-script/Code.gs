/**
 * 食衣住行育樂帳本 — Google Sheets 後端
 *
 * 1. 在試算表開「擴充功能 → Apps Script」，把這整份貼進 Code.gs
 * 2. 上方函式選 setup，按「執行」，授權後到「執行記錄」複製密鑰
 * 3. 「部署 → 新增部署作業 → 網頁應用程式」
 *    執行身分：我；誰可以存取：所有人 → 部署，複製 /exec 網址
 * 4. 在 App 底部「雲端同步設定」貼上網址和密鑰
 *
 * 密鑰存在 Script Properties，不在程式碼裡，所以這份檔案可以公開。
 */

const SHEET = 'Expenses';
const HEAD = ['id', 'date', 'cat', '分類', 'amount', 'note', 'by', 'ts'];
const CAT_NAME = { food: '食', cloth: '衣', home: '住', move: '行', edu: '育', fun: '樂' };

/** 執行一次：建立工作表並產生密鑰。之後再執行只會顯示現有密鑰。 */
function setup() {
  sheet_();
  const props = PropertiesService.getScriptProperties();
  let token = props.getProperty('TOKEN');
  if (!token) {
    token = Utilities.getUuid().replace(/-/g, '');
    props.setProperty('TOKEN', token);
  }
  Logger.log('你的密鑰：' + token);
}

/** 想換一組密鑰時執行（舊裝置要重新貼上新密鑰）。 */
function resetToken() {
  PropertiesService.getScriptProperties().deleteProperty('TOKEN');
  setup();
}

function doGet() {
  return json_({ ok: true, app: 'ledger6' });
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad_request' }); }
  const token = PropertiesService.getScriptProperties().getProperty('TOKEN');
  if (!token || req.token !== token) return json_({ ok: false, error: 'bad_token' });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_();
    if (req.action === 'sync') {
      remove_(sh, req.remove || []);
      upsert_(sh, req.upsert || []);
      return json_({ ok: true, items: read_(sh) });
    }
    return json_({ ok: false, error: 'bad_action' });
  } finally {
    lock.releaseLock();
  }
}

// ---------- helpers

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('B:B').setNumberFormat('@');      // 日期存成文字 YYYY-MM-DD
    sh.getRange('E:E').setNumberFormat('0.00');
  }
  return sh;
}

function rows_(sh) {
  const n = sh.getLastRow() - 1;
  return n > 0 ? sh.getRange(2, 1, n, HEAD.length).getValues() : [];
}

function read_(sh) {
  const tz = Session.getScriptTimeZone();
  return rows_(sh).filter(r => r[0]).map(r => ({
    id: String(r[0]),
    date: r[1] instanceof Date ? Utilities.formatDate(r[1], tz, 'yyyy-MM-dd') : String(r[1]),
    cat: String(r[2]),
    amount: Number(r[4]),
    note: String(r[5] || ''),
    by: String(r[6] || ''),
    ts: Number(r[7]) || 0,
  }));
}

function valid_(it) {
  return it && it.id && CAT_NAME[it.cat] && Number(it.amount) > 0 && /^\d{4}-\d{2}-\d{2}$/.test(String(it.date));
}

function toRow_(it) {
  return [String(it.id), String(it.date), it.cat, CAT_NAME[it.cat], Math.round(Number(it.amount) * 100) / 100,
          String(it.note || '').slice(0, 200), String(it.by || '').slice(0, 40), Number(it.ts) || Date.now()];
}

function upsert_(sh, list) {
  list = list.filter(valid_);
  if (!list.length) return;
  const index = {};
  rows_(sh).forEach((r, i) => { if (r[0]) index[String(r[0])] = i + 2; });
  const append = [];
  list.forEach(it => {
    const at = index[String(it.id)];
    if (at) sh.getRange(at, 1, 1, HEAD.length).setValues([toRow_(it)]);
    else append.push(toRow_(it));
  });
  if (append.length) sh.getRange(sh.getLastRow() + 1, 1, append.length, HEAD.length).setValues(append);
}

function remove_(sh, ids) {
  if (!ids.length) return;
  const kill = new Set(ids.map(String));
  const data = rows_(sh);
  for (let i = data.length - 1; i >= 0; i--) {   // bottom-up so row numbers stay valid
    if (kill.has(String(data[i][0]))) sh.deleteRow(i + 2);
  }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
