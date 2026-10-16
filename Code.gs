/**
 * 子猫お世話タスク API（Google Apps Script / スプレッドシート連動）
 * - スプレッドシートの「logs」シートに完了記録を1行ずつ保存します
 * - 「今日の状態」は logs の今日の日付の行から計算するので、
 *   日付が変われば自動的に全タスクが未完了に戻ります（トリガー不要）
 */

// ★必ず変更してください（夫婦で共有する合言葉）
const PASSCODE = '1234';

const TZ = 'Asia/Tokyo';
// 1日の切り替え時刻。0 = 午前0時。早朝タスクを前日扱いにしたい場合は 4 などに
const DAY_START_HOUR = 0;

const WHO = ['夫', '妻'];
const TASKS = [
  { id: 'poop',    slot: '早朝', label: 'うんち掃除',         icon: '💩' },
  { id: 'food_am', slot: '朝',   label: 'フード',             icon: '🍚' },
  { id: 'food_pm', slot: '夕方', label: 'フード',             icon: '🍚' },
  { id: 'water',   slot: '夕方', label: '水交換',             icon: '💧' },
  { id: 'pee',     slot: '夜',   label: 'おしっこシート交換', icon: '🧻' },
];

/* ---------- エントリーポイント ---------- */
function doGet(e) {
  return handle_(e.parameter || {});
}

function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) {}
  return handle_(body);
}

function handle_(p) {
  let res;
  try {
    if (String(p.pass) !== PASSCODE) {
      res = { ok: false, error: 'auth' };
    } else if (p.action === 'state') {
      res = Object.assign({ ok: true }, getState_());
    } else if (p.action === 'history') {
      res = getHistory_(Number(p.days) || 14);
    } else if (p.action === 'complete') {
      res = complete_(p.taskId, p.who);
    } else if (p.action === 'undo') {
      res = undo_(p.taskId);
    } else {
      res = { ok: false, error: 'unknown_action' };
    }
  } catch (err) {
    res = { ok: false, error: String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(res))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ---------- シート操作 ---------- */
function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('logs');
  if (!sh) {
    sh = ss.insertSheet('logs');
    sh.appendRow(['date', 'taskId', 'who', 'time', 'timestamp']);
    sh.getRange('A:D').setNumberFormat('@'); // 日付・時刻を文字列のまま保存
    sh.setFrozenRows(1);
  }
  return sh;
}

function fmtDate_(v) { return v instanceof Date ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd') : String(v); }
function fmtTime_(v) { return v instanceof Date ? Utilities.formatDate(v, TZ, 'HH:mm') : String(v); }

function rows_() {
  const sh = sheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 4).getValues().map(function (r, i) {
    return { row: i + 2, date: fmtDate_(r[0]), taskId: String(r[1]), who: String(r[2]), time: fmtTime_(r[3]) };
  });
}

function dayString_(offsetDays) {
  const ms = Date.now() - DAY_START_HOUR * 3600 * 1000 - (offsetDays || 0) * 86400 * 1000;
  return Utilities.formatDate(new Date(ms), TZ, 'yyyy-MM-dd');
}

/* ---------- 今日の状態 ---------- */
function getState_() {
  const today = dayString_(0);
  const todays = rows_().filter(function (r) { return r.date === today; });
  const tasks = TASKS.map(function (t) {
    const hit = todays.filter(function (r) { return r.taskId === t.id; })[0];
    return Object.assign({}, t, { done: !!hit, who: hit ? hit.who : null, time: hit ? hit.time : null });
  });
  return { today: today, tasks: tasks };
}

/* ---------- 完了（二重実行防止つき） ---------- */
function complete_(taskId, who) {
  if (!TASKS.some(function (t) { return t.id === taskId; })) return { ok: false, error: 'bad_task' };
  if (WHO.indexOf(who) < 0) return { ok: false, error: 'bad_who' };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000); // 同時タップでも1件だけ通す
  try {
    const today = dayString_(0);
    const dup = rows_().filter(function (r) { return r.date === today && r.taskId === taskId; })[0];
    if (dup) {
      return Object.assign({ ok: false, duplicate: true, taskId: taskId, who: dup.who, time: dup.time }, getState_());
    }
    const now = new Date();
    sheet_().appendRow([today, taskId, who, Utilities.formatDate(now, TZ, 'HH:mm'), now.toISOString()]);
    return Object.assign({ ok: true }, getState_());
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 取り消し（押し間違い用・今日の分のみ） ---------- */
function undo_(taskId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const today = dayString_(0);
    const rows = rows_().filter(function (r) { return r.date === today && r.taskId === taskId; });
    rows.reverse().forEach(function (r) { sheet_().deleteRow(r.row); });
    return Object.assign({ ok: true }, getState_());
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 履歴 ---------- */
function getHistory_(days) {
  days = Math.min(Math.max(days, 1), 90);
  const all = rows_();
  const out = [];
  for (let i = 0; i < days; i++) {
    const d = dayString_(i);
    const entries = {};
    all.filter(function (r) { return r.date === d; })
       .forEach(function (r) { entries[r.taskId] = { who: r.who, time: r.time }; });
    out.push({ date: d, entries: entries });
  }
  return { ok: true, tasks: TASKS, days: out };
}
