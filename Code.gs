/**
 * 裝潢甘特圖 — Google Sheets API（Apps Script）
 *
 * 這支程式綁在你的 Google 試算表上，只負責「讀寫資料」，不顯示任何畫面。
 * 甘特圖網頁放在 GitHub Pages，透過這裡的 doPost 存取資料。
 *
 * 密碼「不寫在程式碼裡」：請到 Apps Script 左側「專案設定 ⚙️」→「指令碼屬性」
 * 新增屬性名稱 PASSWORD，值填你們要用的密碼。
 */

const SHEET_NAME = '時程';
const HEADERS = ['id', '工項', '開始', '結束', '進度', '相依', '工班', '備註', '類型'];
const COL = { id: 1, name: 2, start: 3, end: 4, progress: 5, deps: 6, crew: 7, note: 8, type: 9 };
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
      ['T1', '拆除工程', d(0), d(3), 0, '', '拆除班', '', ''],
      ['T2', '水電配管', d(4), d(10), 0, 'T1', '水電', '', ''],
      ['T3', '泥作',     d(11), d(20), 0, 'T2', '泥作', '', ''],
      ['T4', '木作',     d(21), d(35), 0, 'T3', '木工', '範例資料，可刪除', ''],
    ]);
  }

  const pw = PropertiesService.getScriptProperties().getProperty('PASSWORD');
  Logger.log(pw
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
 *   { password, action: 'list' | 'update' | 'add' | 'delete', ... }
 * 密碼放在 body，不放在網址上。
 */
function doPost(e) {
  let req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: '格式錯誤' });
  }

  if (!checkPassword_(req.password)) {
    Utilities.sleep(1500); // 拖慢猜密碼的速度
    return json_({ ok: false, error: 'AUTH' });
  }

  try {
    switch (req.action) {
      case 'list':   return json_({ ok: true, tasks: listTasks_() });
      case 'update': return json_({ ok: true, task: withLock_(() => updateTask_(req.id, req.fields)) });
      case 'add':    return json_({ ok: true, task: withLock_(() => addTask_(req.fields)) });
      case 'delete': return json_({ ok: true, id: withLock_(() => deleteTask_(req.id)) });
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
  };
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

function checkPassword_(input) {
  const pw = PropertiesService.getScriptProperties().getProperty('PASSWORD');
  return !!pw && typeof input === 'string' && input === pw;
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
