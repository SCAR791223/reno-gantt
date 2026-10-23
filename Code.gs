/**
 * 裝潢甘特圖 — Google Sheets API（Apps Script）
 *
 * 這支程式綁在你的 Google 試算表上，只負責「讀寫資料」，不顯示任何畫面。
 * 甘特圖網頁放在 GitHub Pages，透過這裡的 doPost 存取資料。
 *
 * 密碼「不寫在程式碼裡」：請到 Apps Script 左側「專案設定 ⚙️」→「指令碼屬性」新增：
 *   PASSWORD   可以看也可以改（你和老婆用）
 * 瀏覽連結的鑰匙 VIEW_KEY 不用自己設：在網頁按「分享瀏覽連結」時自動產生。
 *
 * LINE 群組通知（選用）：
 *   LINE_TOKEN  LINE Messaging API 的 Channel access token（自己貼上）
 *   LINE_GROUP  要通知的群組，在網頁按「🔔 LINE 通知 → 開始連結」後把機器人加進群組就會自動記下
 *   每天早上 8 點由 setupNotify() 建立的排程檢查一次，有事才發。
 */

// 提前幾天提醒
const NOTIFY_MS_DAYS = [7, 3];     // 注意事項：截止前 7 天、3 天
const NOTIFY_TASK_DAYS = [3];      // 工程：開工前 3 天

const SHEET_NAME = '時程';
const HEADERS = ['id', '工項', '開始', '結束', '進度', '相依', '工班', '備註', '類型', '排序'];
const COL = { id: 1, name: 2, start: 3, end: 4, progress: 5, deps: 6, crew: 7, note: 8, type: 9, order: 10 };
// 「排序」欄：網頁上拖拉清單決定的順序；空白的照開始日期排在後面
// 「類型」欄：空白 = 一般工項；「截止」= 注意事項（菱形）
const TYPE_MS = '截止';
const isMsCell_ = (v) => /截止|注意|里程碑|milestone/i.test(String(v || ''));

/* ───────────────────────── 第一次使用：手動執行一次 ───────────────────────── */

/**
 * 建立「時程」工作表、標題列與日期格式。
 * 在 Apps Script 編輯器上方選擇 setup，按「執行」。第一次會要求授權。
 */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);

  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
    .setFontWeight('bold').setBackground('#f1f3f4');
  sh.setFrozenRows(1);
  sh.getRange('C:D').setNumberFormat('yyyy-mm-dd');
  sh.getRange('E:E').setNumberFormat('0');
  sh.setColumnWidth(COL.name, 180);
  sh.setColumnWidth(COL.note, 240);

  // 只有空表時放幾筆範例，方便先看效果；之後直接刪掉或改掉即可
  if (sh.getLastRow() < 2) {
    const today = new Date();
    const d = (offset) => fmtDate_(new Date(today.getTime() + offset * 86400000), ss);
    sh.getRange(2, 1, 4, HEADERS.length).setValues([
      ['T1', '拆除工程', d(0), d(3), 0, '', '拆除', '', '', ''],
      ['T2', '水電配管', d(4), d(10), 0, 'T1', '水電', '', '', ''],
      ['T3', '泥作',     d(11), d(20), 0, 'T2', '泥作', '', '', ''],
      ['T4', '木作',     d(21), d(35), 0, 'T3', '木工', '範例資料，可刪除', '', ''],
    ]);
  }

  const props = PropertiesService.getScriptProperties();
  Logger.log(props.getProperty('PASSWORD')
    ? '✅ 設定完成，密碼已設定。'
    : '⚠️ 設定完成，但還沒設密碼：請到「專案設定 → 指令碼屬性」新增 PASSWORD。');
}

/* ───────────────────────── API 入口 ───────────────────────── */

/** 直接用瀏覽器打開網址時顯示，不回傳任何資料 */
function doGet() {
  return json_({ ok: false, error: 'Use POST' });
}

/**
 * 所有請求都用 POST，body 是 JSON 文字：
 *   { password?, key?, action: 'list' | 'update' | 'add' | 'delete' | 'reorder' | 'shareKey' | 'line…', ... }
 * password = 修改密碼；key = 分享連結裡的瀏覽鑰匙（只能 list）。
 */
function doPost(e) {
  let req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: '格式錯誤' });
  }

  // LINE 傳來的 webhook（機器人被加進群組、群組裡有人傳訊息）
  if (req && Array.isArray(req.events)) {
    try { handleLine_(req.events); } catch (err) { console.error(err); }
    return json_({ ok: true });
  }

  const role = roleOf_(req.password, req.key);
  if (!role) {
    Utilities.sleep(1500); // 拖慢猜密碼的速度
    return json_({ ok: false, error: 'AUTH' });
  }
  if (req.action !== 'list' && role !== 'edit') {
    return json_({ ok: false, error: 'READONLY' });
  }

  try {
    switch (req.action) {
      case 'list':   return json_({ ok: true, role, tasks: listTasks_() });
      case 'shareKey':
        if (req.site) PropertiesService.getScriptProperties().setProperty('SITE_URL', String(req.site).slice(0, 200));
        return json_({ ok: true, key: shareKey_(!!req.renew) });
      case 'update': return json_({ ok: true, task: withLock_(() => updateTask_(req.id, req.fields)) });
      case 'add':    return json_({ ok: true, task: withLock_(() => addTask_(req.fields)) });
      case 'delete': return json_({ ok: true, id: withLock_(() => deleteTask_(req.id)) });
      case 'reorder': return json_({ ok: true, n: withLock_(() => reorder_(req.ids)) });
      case 'lineStatus': return json_({ ok: true, line: lineStatus_() });
      case 'lineBind':   return json_({ ok: true, line: lineBind_() });
      case 'lineUnbind': return json_({ ok: true, line: lineUnbind_() });
      case 'lineTest':   return json_({ ok: true, sent: lineTest_() });
      default:       return json_({ ok: false, error: '未知的動作' });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

/* ───────────────────────── 資料操作 ───────────────────────── */

function listTasks_() {
  const sh = sheet_();
  const ss = sh.getParent();
  const last = sh.getLastRow();
  if (last < 2) return [];

  const values = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
  const used = new Set(values.map(r => String(r[0]).trim()).filter(Boolean));
  const tasks = [];
  let assigned = false;

  values.forEach((r, i) => {
    const name = String(r[COL.name - 1]).trim();
    if (!name) return; // 空白列略過

    // 在 Sheets 手動新增的列沒有 id，這裡自動補上
    let id = String(r[COL.id - 1]).trim();
    if (!id) {
      id = nextId_(used);
      used.add(id);
      sh.getRange(i + 2, COL.id).setValue(id);
      assigned = true;
    }

    const start = toDateStr_(r[COL.start - 1], ss);
    let end = toDateStr_(r[COL.end - 1], ss) || start;
    if (!start) return; // 沒有開始日無法畫，略過
    if (end < start) end = start;

    tasks.push({
      id,
      name,
      start,
      end,
      progress: clamp_(Number(r[COL.progress - 1]) || 0, 0, 100),
      dependencies: String(r[COL.deps - 1] || '').split(/[,，、\s]+/).filter(Boolean).join(','),
      crew: String(r[COL.crew - 1] || '').trim(),
      note: String(r[COL.note - 1] || '').trim(),
      type: isMsCell_(r[COL.type - 1]) ? 'milestone' : 'task',
      order: orderOf_(r[COL.order - 1]),
    });
  });

  if (assigned) SpreadsheetApp.flush();
  return tasks;
}

function updateTask_(id, fields) {
  const sh = sheet_();
  const row = findRow_(sh, id);
  if (!row) throw new Error('找不到這個工項，可能已被刪除，請重新整理');
  writeFields_(sh, row, fields || {});
  return readRow_(sh, row);
}

function addTask_(fields) {
  if (!fields || !String(fields.name || '').trim()) throw new Error('工項名稱不能空白');
  if (!fields.start) throw new Error('請填開始日');
  const sh = sheet_();
  const ids = new Set(sh.getLastRow() > 1
    ? sh.getRange(2, COL.id, sh.getLastRow() - 1, 1).getValues().map(r => String(r[0]).trim())
    : []);
  const id = nextId_(ids);
  sh.appendRow([id]);
  const row = sh.getLastRow();
  writeFields_(sh, row, Object.assign({ progress: 0, end: fields.start }, fields));
  // 已經手動排過順序的話，新項目排在最後
  const orders = row > 2 ? sh.getRange(2, COL.order, row - 2, 1).getValues().map(([v]) => orderOf_(v)).filter((v) => v !== null) : [];
  if (orders.length) sh.getRange(row, COL.order).setValue(Math.max.apply(null, orders) + 1);
  return readRow_(sh, row);
}

function deleteTask_(id) {
  const sh = sheet_();
  const row = findRow_(sh, id);
  if (!row) return id;
  sh.deleteRow(row);

  // 把其他工項「相依」欄裡的這個 id 一起拿掉
  const last = sh.getLastRow();
  if (last >= 2) {
    const rng = sh.getRange(2, COL.deps, last - 1, 1);
    const deps = rng.getValues().map(([v]) => [
      String(v || '').split(/[,，、\s]+/).filter(x => x && x !== id).join(','),
    ]);
    rng.setValues(deps);
  }
  return id;
}

/** 依網頁拖拉後的順序，把 1、2、3… 寫進「排序」欄 */
function reorder_(ids) {
  if (!Array.isArray(ids) || !ids.length) throw new Error('沒有順序資料');
  const sh = sheet_();
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const pos = {};
  ids.forEach((id, i) => { pos[String(id)] = i + 1; });
  const rows = sh.getRange(2, COL.id, last - 1, 1).getValues();
  const rng = sh.getRange(2, COL.order, last - 1, 1);
  const cur = rng.getValues();
  // 沒在清單裡的（例如網頁開著時有人在 Sheets 新增的列）排在最後
  let extra = ids.length;
  const out = rows.map(([id], i) => {
    const k = String(id).trim();
    if (k in pos) return [pos[k]];
    return [k ? ++extra : cur[i][0]];
  });
  rng.setValues(out);
  return ids.length;
}

/* ───────────────────────── LINE 群組通知 ───────────────────────── */

/** 第一次使用：在編輯器選 setupNotify 按「執行」，授權後每天早上 8 點自動檢查 */
function setupNotify() {
  ScriptApp.getProjectTriggers()
    .filter((t) => t.getHandlerFunction() === 'dailyNotify')
    .forEach((t) => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('dailyNotify').timeBased().atHour(8).everyDays(1)
    .inTimezone(SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone()).create();
  UrlFetchApp.fetch('https://api.line.me/v2/bot/info', { muteHttpExceptions: true }); // 順便取得連網授權
  Logger.log('✅ 已設定：每天早上 8 點檢查要不要發 LINE 通知');
}

/** 排程每天呼叫：有要提醒的事才發到群組 */
function dailyNotify() {
  const p = PropertiesService.getScriptProperties();
  if (!p.getProperty('LINE_TOKEN') || !p.getProperty('LINE_GROUP')) return;
  const text = buildDigest_(false);
  if (text) linePush_(p.getProperty('LINE_GROUP'), text);
}

/** 組出今天的提醒內容；force=true 時沒有事也回一段話（測試用） */
function buildDigest_(force) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const today = fmtDate_(new Date(), ss);
  const tasks = listTasks_();
  const ms = [], works = [];
  tasks.forEach((t) => {
    const d = daysBetween_(today, t.start);
    if (t.type === 'milestone') {
      if (t.progress < 100 && NOTIFY_MS_DAYS.indexOf(d) >= 0) ms.push({ t, d });
    } else if (t.progress < 100 && NOTIFY_TASK_DAYS.indexOf(d) >= 0) works.push({ t, d });
  });
  if (!ms.length && !works.length && !force) return '';
  const line = ({ t, d }, tail) => `・${d} 天後 ${md_(t.start)}（${wd_(t.start)}）${t.name}${tail}${t.crew ? `（${t.crew}）` : ''}`;
  const out = [`🏠 裝潢提醒 ${md_(today)}（${wd_(today)}）`];
  if (ms.length) out.push('', '◆ 注意事項快到期', ...ms.sort((a, b) => a.d - b.d).map((x) => line(x, ' 截止')));
  if (works.length) out.push('', '● 工程快開工', ...works.sort((a, b) => a.d - b.d).map((x) => line(x, ' 開工')));
  if (!ms.length && !works.length) {
    const next = tasks.filter((t) => t.start >= today && t.progress < 100).sort((a, b) => a.start.localeCompare(b.start))[0];
    out.push('', '這幾天沒有需要提醒的事。', next ? `下一件：${md_(next.start)}（${wd_(next.start)}）${next.name}` : '');
  }
  const key = PropertiesService.getScriptProperties().getProperty('VIEW_KEY');
  const site = PropertiesService.getScriptProperties().getProperty('SITE_URL') || 'https://scar791223.github.io/reno-gantt/';
  if (key && site) out.push('', `📅 完整時程：${site}#v=${key}`);
  return out.join('\n').trim();
}

function handleLine_(events) {
  const p = PropertiesService.getScriptProperties();
  const bound = p.getProperty('LINE_GROUP');
  const until = Number(p.getProperty('LINE_BIND_UNTIL') || 0);
  events.forEach((ev) => {
    const src = ev.source || {};
    const gid = src.groupId || src.roomId;
    if (!gid) return;
    const text = ev.type === 'message' && ev.message && ev.message.type === 'text' ? String(ev.message.text).trim() : '';
    // 只有在網頁按了「開始連結」後 10 分鐘內，才接受新的群組（避免別人亂綁）
    if (Date.now() < until && (ev.type === 'join' || /綁定|連結/.test(text))) {
      p.setProperty('LINE_GROUP', gid);
      p.deleteProperty('LINE_BIND_UNTIL');
      lineReply_(ev.replyToken, '✅ 已連結裝潢時程通知\n注意事項截止前 7 天、3 天，工程開工前 3 天，早上 8 點會在這裡提醒。\n想看今天的提醒，傳「提醒」就好。');
      return;
    }
    if (gid === bound && /^(提醒|通知|時程)$/.test(text)) {
      lineReply_(ev.replyToken, buildDigest_(true));
    }
  });
}

function lineStatus_() {
  const p = PropertiesService.getScriptProperties();
  const until = Number(p.getProperty('LINE_BIND_UNTIL') || 0);
  return {
    hasToken: !!p.getProperty('LINE_TOKEN'),
    bound: !!p.getProperty('LINE_GROUP'),
    binding: Date.now() < until,
    hasTrigger: ScriptApp.getProjectTriggers().some((t) => t.getHandlerFunction() === 'dailyNotify'),
  };
}
function lineBind_() {
  PropertiesService.getScriptProperties().setProperty('LINE_BIND_UNTIL', String(Date.now() + 10 * 60 * 1000));
  return lineStatus_();
}
function lineUnbind_() {
  const p = PropertiesService.getScriptProperties();
  p.deleteProperty('LINE_GROUP'); p.deleteProperty('LINE_BIND_UNTIL');
  return lineStatus_();
}
function lineTest_() {
  const g = PropertiesService.getScriptProperties().getProperty('LINE_GROUP');
  if (!g) throw new Error('還沒連結 LINE 群組');
  linePush_(g, buildDigest_(true));
  return true;
}

function linePush_(to, text) { lineCall_('push', { to, messages: [{ type: 'text', text: text.slice(0, 4900) }] }); }
function lineReply_(token, text) { if (token) lineCall_('reply', { replyToken: token, messages: [{ type: 'text', text: text.slice(0, 4900) }] }); }
function lineCall_(kind, body) {
  const token = PropertiesService.getScriptProperties().getProperty('LINE_TOKEN');
  if (!token) throw new Error('還沒設定 LINE_TOKEN');
  const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/' + kind, {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + token }, payload: JSON.stringify(body),
  });
  if (res.getResponseCode() >= 300) throw new Error('LINE 回應 ' + res.getResponseCode() + '：' + res.getContentText().slice(0, 200));
}

const WD_ = ['日', '一', '二', '三', '四', '五', '六'];
function noon_(s) { return new Date(s + 'T12:00:00Z'); }
function daysBetween_(a, b) { return Math.round((noon_(b) - noon_(a)) / 86400000); }
function md_(s) { return Number(s.slice(5, 7)) + '/' + Number(s.slice(8, 10)); }
function wd_(s) { return WD_[noon_(s).getUTCDay()]; }

/* ───────────────────────── 工具函式 ───────────────────────── */

function writeFields_(sh, row, f) {
  const set = (col, v) => sh.getRange(row, col).setValue(v);
  if ('name' in f)         set(COL.name, String(f.name).trim());
  if ('start' in f)        set(COL.start, validDate_(f.start));
  if ('end' in f)          set(COL.end, validDate_(f.end));
  if ('progress' in f)     set(COL.progress, Math.round(clamp_(Number(f.progress) || 0, 0, 100)));
  if ('dependencies' in f) set(COL.deps, String(f.dependencies || ''));
  if ('crew' in f)         set(COL.crew, String(f.crew || '').trim());
  if ('note' in f)         set(COL.note, String(f.note || '').trim());
  if ('type' in f)         set(COL.type, f.type === 'milestone' ? TYPE_MS : '');

  // 確保結束日不早於開始日
  const ss = sh.getParent();
  const s = toDateStr_(sh.getRange(row, COL.start).getValue(), ss);
  const e = toDateStr_(sh.getRange(row, COL.end).getValue(), ss);
  if (s && (!e || e < s)) set(COL.end, s);
}

function readRow_(sh, row) {
  const ss = sh.getParent();
  const r = sh.getRange(row, 1, 1, HEADERS.length).getValues()[0];
  return {
    id: String(r[0]),
    name: String(r[1]),
    start: toDateStr_(r[2], ss),
    end: toDateStr_(r[3], ss),
    progress: Number(r[4]) || 0,
    dependencies: String(r[5] || ''),
    crew: String(r[6] || ''),
    note: String(r[7] || ''),
    type: isMsCell_(r[8]) ? 'milestone' : 'task',
    order: orderOf_(r[9]),
  };
}

function orderOf_(v) {
  if (v === '' || v === null || v === undefined) return null;
  const n = Number(v);
  return isFinite(n) ? n : null;
}

function sheet_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('找不到「' + SHEET_NAME + '」工作表，請先執行 setup');
  // 舊版試算表沒有新加的欄位，自動補上標題
  if (sh.getLastColumn() < HEADERS.length) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
      .setFontWeight('bold').setBackground('#f1f3f4');
  }
  return sh;
}

function findRow_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2 || !id) return 0;
  const ids = sh.getRange(2, COL.id, last - 1, 1).getValues();
  const i = ids.findIndex(([v]) => String(v).trim() === String(id));
  return i < 0 ? 0 : i + 2;
}

function nextId_(used) {
  let n = 1;
  used.forEach(id => {
    const m = /^T(\d+)$/.exec(id);
    if (m) n = Math.max(n, Number(m[1]) + 1);
  });
  while (used.has('T' + n)) n++;
  return 'T' + n;
}

/** Sheets 裡的日期可能是 Date 物件或文字，統一轉成 yyyy-MM-dd */
function toDateStr_(v, ss) {
  if (!v) return '';
  if (v instanceof Date) return fmtDate_(v, ss);
  const m = /^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})/.exec(String(v).trim());
  return m ? `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}` : '';
}

function fmtDate_(d, ss) {
  return Utilities.formatDate(d, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
}

function validDate_(s) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(s))) throw new Error('日期格式應為 yyyy-mm-dd');
  return String(s);
}

function clamp_(n, lo, hi) { return Math.min(hi, Math.max(lo, n)); }

/** 回傳 'edit'（可修改）、'view'（只能看）或 null（都不對） */
function roleOf_(password, key) {
  const props = PropertiesService.getScriptProperties();
  const pw = props.getProperty('PASSWORD');
  const viewKey = props.getProperty('VIEW_KEY');
  if (pw && typeof password === 'string' && password && password === pw) return 'edit';
  if (viewKey && typeof key === 'string' && key && key === viewKey) return 'view';
  return null;
}

/** 取得分享連結用的瀏覽鑰匙；沒有或要求換新時產生新的（舊連結隨即失效） */
function shareKey_(renew) {
  const props = PropertiesService.getScriptProperties();
  let key = props.getProperty('VIEW_KEY');
  if (!key || renew) {
    key = Utilities.getUuid().replace(/-/g, '').slice(0, 20);
    props.setProperty('VIEW_KEY', key);
  }
  return key;
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000); // 兩人同時改時排隊處理
  try { return fn(); } finally { lock.releaseLock(); }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
