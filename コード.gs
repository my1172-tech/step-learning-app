/**
 * ステップ学習アプリ 集約・課題配信用 Google Apps Script
 * スプレッドシートに紐づけて使います（拡張機能 → Apps Script）。
 *
 * シート（初回に setup() を実行すると自動で作られます）
 *   名簿 : ID | 学年 | 組 | 出席番号 | パスワード   ※氏名は保存しない
 *   教員 : 教員ID | 氏名 | パスワード
 *   課題 : 課題ID | 課題名 | 教科 | 単元 | 問題ファイル | 対象学年 | 対象組 |
 *          公開開始 | 締切 | 出題数 | 出題順 | 再挑戦 | 担当 | 有効 | 登録日時
 *   回答 : 送信日時 | バッチID | 課題ID | 課題名 | 時間外 | ID | 学年 | 組 | 出席番号 |
 *          問題ID | 学年(問題) | 教科 | 単元 | STEP | 問題文 |
 *          えらんだ答え | えらんだ説明 | 正誤 | 正解数 | 設問数 | 所要秒
 */

/* ====== 設定 ====== */
const TOKEN  = "××××××";               // index.html / teacher.html の SEND_TOKEN と同じ
const SECRET = "hakone-2026-taku-secret";  // ログイン券の生成用。長い文字列にして外部に出さない
const TZ     = "Asia/Tokyo";
const QUESTION_FOLDER_ID = "";             // 問題CSVを置くDriveフォルダのID（URLの /folders/ の後ろ）
const LATE_DAYS = 7;                       // 締切後、何日まで「時間外」として受け付けるか

const SH_ANSWER = "回答", SH_ROSTER = "名簿", SH_TEACHER = "教員", SH_TASK = "課題";

const ANSWER_HEADER = ["送信日時","バッチID","課題ID","課題名","時間外",
  "ID","学年","組","出席番号",
  "問題ID","学年(問題)","教科","単元","STEP","問題文",
  "えらんだ答え","えらんだ説明","正誤","正解数","設問数","所要秒"];
const ROSTER_HEADER  = ["ID","学年","組","出席番号","パスワード"];  // 氏名は保存しない
const TEACHER_HEADER = ["教員ID","氏名","パスワード"];
const TASK_HEADER    = ["課題ID","課題名","教科","単元","問題ファイル","対象学年","対象組",
  "公開開始","締切","出題数","出題順","再挑戦","担当","有効","登録日時"];

/* ====== セットアップ ====== */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet_(ss, SH_ROSTER, ROSTER_HEADER);
  ensureSheet_(ss, SH_TEACHER, TEACHER_HEADER);
  ensureSheet_(ss, SH_TASK, TASK_HEADER);
  ensureSheet_(ss, SH_ANSWER, ANSWER_HEADER);
  SpreadsheetApp.getUi().alert("シートを用意しました。名簿と教員シートに登録してください。");
}

function パスワードをハッシュ化() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let n = 0;
  [[SH_ROSTER, 5], [SH_TEACHER, 3]].forEach(function (t) {
    const sh = ss.getSheetByName(t[0]);
    if (!sh || sh.getLastRow() < 2) return;
    const rng = sh.getRange(2, t[1], sh.getLastRow() - 1, 1);
    const v = rng.getValues();
    for (let i = 0; i < v.length; i++) {
      const s = String(v[i][0] || "").trim();
      if (s && !isHash_(s)) { v[i][0] = sha256_(s); n++; }
    }
    rng.setValues(v);
  });
  SpreadsheetApp.getUi().alert(n + "件をハッシュに置き換えました。");
}

function ensureSheet_(ss, name, header) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight("bold");
    sh.setFrozenRows(1);
  }
  return sh;
}

/* ====== POST ====== */
function doPost(e) {
  let p;
  try { p = JSON.parse(e.postData.contents); } catch (err) { return textOut_("NG:json"); }
  if (TOKEN && p.token !== TOKEN) return textOut_("NG:token");
  if (p.action === "login")    return handleLogin_(p);
  if (p.action === "saveTask") return handleSaveTask_(p);
  if (p.action === "submit")   return handleSubmit_(p);
  return textOut_("NG:action");
}

function handleLogin_(p) {
  const role = (p.role === "teacher") ? "teacher" : "student";
  const id = String(p.id || "").trim();
  const cache = CacheService.getScriptCache();
  const key = "login:" + p.nonce;
  const failKey = "fail:" + role + ":" + id;

  if (Number(cache.get(failKey) || 0) >= 5) {
    cache.put(key, JSON.stringify({ ok: false, msg: "しばらく待ってからやり直してください。" }), 180);
    return textOut_("OK");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const who = (role === "teacher") ? lookupTeacher_(ss, id) : lookupRoster_(ss, id);
  const stored = String(who.pw || "").trim();
  const ok = who.found && stored &&
    (isHash_(stored) ? stored.toLowerCase() : sha256_(stored)) === String(p.pwHash || "").toLowerCase();

  if (!ok) {
    cache.put(failKey, String(Number(cache.get(failKey) || 0) + 1), 600);
    cache.put(key, JSON.stringify({ ok: false, msg: "IDかパスワードがちがいます。" }), 180);
    return textOut_("OK");
  }
  cache.remove(failKey);
  cache.put(key, JSON.stringify({
    ok: true, role: role,
    who: (role === "teacher") ? String(who.name) + " 先生"
                              : who.grade + "年" + who.cls + "組" + who.no + "番",
    ticket: ticketFor_(role, id, todayStr_())
  }), 180);
  return textOut_("OK");
}

function handleSubmit_(p) {
  const id = String(p.id || "").trim();
  if (!id || !p.rows || !p.rows.length) return textOut_("NG:empty");
  if (!verifyTicket_("student", id, p.ticket)) return textOut_("NG:ticket");

  const lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch (err) { return textOut_("NG:busy"); }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ensureSheet_(ss, SH_ANSWER, ANSWER_HEADER);
    const who = lookupRoster_(ss, id);
    const task = p.taskId ? findTask_(ss, p.taskId) : null;
    const now = new Date();
    /* 締切判定は端末時計ではなくサーバ時刻で行う */
    const late = (task && task.due && now.getTime() > task.due.getTime()) ? "時間外" : "";
    const head = [now, p.batch, (task ? task.id : ""), (task ? task.name : ""), late,
                  id, who.grade, who.cls, who.no];
    const rows = p.rows.map(function (r) { return head.concat(r); });
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, ANSWER_HEADER.length).setValues(rows);
    return textOut_("OK");
  } catch (err) {
    return textOut_("NG:" + err);
  } finally { lock.releaseLock(); }
}

function handleSaveTask_(p) {
  const id = String(p.id || "").trim();
  if (!verifyTicket_("teacher", id, p.ticket)) return textOut_("NG:ticket");
  const t = p.task || {};
  if (!t.name || !t.file) return textOut_("NG:empty");

  const lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch (err) { return textOut_("NG:busy"); }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ensureSheet_(ss, SH_TASK, TASK_HEADER);
    const row = [
      t.id || newTaskId_(sh), t.name, t.subject || "", t.unit || "", t.file,
      t.grade || "", t.cls || "", t.open || "", t.due || "",
      t.count || "", t.order || "ランダム", t.retry || "可",
      String(p.teacherName || id), t.active || "○", new Date()
    ];
    /* 課題IDが既にあれば上書き、なければ追加 */
    const at = t.id ? taskRowIndex_(sh, t.id) : 0;
    if (at) sh.getRange(at, 1, 1, TASK_HEADER.length).setValues([row]);
    else    sh.getRange(sh.getLastRow() + 1, 1, 1, TASK_HEADER.length).setValues([row]);
    return textOut_("OK");
  } catch (err) {
    return textOut_("NG:" + err);
  } finally { lock.releaseLock(); }
}

/* ====== GET（JSONP） ====== */
function doGet(e) {
  const a = e.parameter.action;
  let out = { ok: false };

  if (a === "loginResult" && e.parameter.nonce) {
    const cache = CacheService.getScriptCache();
    const k = "login:" + e.parameter.nonce;
    const hit = cache.get(k);
    if (hit) { cache.remove(k); out = JSON.parse(hit); }
    else out = { ok: false, msg: "ログインの確認ができませんでした。もう一度おしてください。" };

  } else if (a === "verify" && e.parameter.batch) {
    out = { ok: false, count: 0 };
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_ANSWER);
    if (sh && sh.getLastRow() > 1) {
      const col = sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues();
      const b = String(e.parameter.batch);
      out.count = col.filter(function (v) { return String(v[0]) === b; }).length;
      out.ok = out.count > 0;
    }

  } else if (a === "tasks") {
    out = verifyTicket_("student", e.parameter.id, e.parameter.ticket)
      ? { ok: true, tasks: studentTasks_(String(e.parameter.id).trim()) }
      : { ok: false, msg: "ログインし直してください。" };

  } else if (a === "questions") {
    out = verifyTicket_("student", e.parameter.id, e.parameter.ticket)
      ? questionsFor_(e.parameter.taskId)
      : { ok: false, msg: "ログインし直してください。" };

  } else if (a === "teacherTasks") {
    out = verifyTicket_("teacher", e.parameter.id, e.parameter.ticket)
      ? { ok: true, tasks: allTasks_(), files: csvFiles_() }
      : { ok: false, msg: "ログインし直してください。" };
  }

  const json = JSON.stringify(out);
  const cb = e.parameter.callback;
  if (cb && /^[A-Za-z0-9_]+$/.test(cb)) {
    return ContentService.createTextOutput(cb + "(" + json + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

/* ====== 課題 ====== */
function taskSheetValues_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_TASK);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, TASK_HEADER.length).getValues();
}
function toTask_(r) {
  return { id: String(r[0]), name: String(r[1]), subject: String(r[2]), unit: String(r[3]),
    file: String(r[4]), grade: String(r[5]).trim(), cls: String(r[6]).trim(),
    open: toDate_(r[7]), due: toDate_(r[8]),
    count: r[9], order: String(r[10] || "ランダム"), retry: String(r[11] || "可"),
    owner: String(r[12]), active: String(r[13]).trim() };
}
function toDate_(v) {
  if (!v) return null;
  if (Object.prototype.toString.call(v) === "[object Date]") return v;
  const d = new Date(String(v).replace(/-/g, "/"));
  return isNaN(d.getTime()) ? null : d;
}
function findTask_(ss, id) {
  const vals = taskSheetValues_();
  for (let i = 0; i < vals.length; i++) if (String(vals[i][0]) === String(id)) return toTask_(vals[i]);
  return null;
}
function taskRowIndex_(sh, id) {
  if (sh.getLastRow() < 2) return 0;
  const col = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < col.length; i++) if (String(col[i][0]) === String(id)) return i + 2;
  return 0;
}
function newTaskId_(sh) {
  return "T" + todayStr_() + "-" + (sh.getLastRow());
}

/** 生徒に見せる課題（対象・公開期間・再挑戦可否で絞る） */
function studentTasks_(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const me = lookupRoster_(ss, id);
  const now = new Date();
  const done = submittedTaskIds_(ss, id);
  const out = [];
  taskSheetValues_().forEach(function (r) {
    const t = toTask_(r);
    if (t.active !== "○") return;
    if (t.open && now.getTime() < t.open.getTime()) return;
    if (t.due && now.getTime() > t.due.getTime() + LATE_DAYS * 86400000) return;
    if (t.grade && String(t.grade) !== String(me.grade)) return;
    if (t.cls && t.cls.split(/[,、]/).map(function (s) { return s.trim(); }).indexOf(String(me.cls)) < 0) return;
    const submitted = done.indexOf(t.id) >= 0;
    if (submitted && t.retry !== "可") return;
    out.push({ id: t.id, name: t.name, subject: t.subject, unit: t.unit,
      due: t.due ? Utilities.formatDate(t.due, TZ, "M月d日 HH:mm") : "",
      late: !!(t.due && now.getTime() > t.due.getTime()),
      count: t.count, order: t.order, submitted: submitted });
  });
  return out;
}

function submittedTaskIds_(ss, id) {
  const sh = ss.getSheetByName(SH_ANSWER);
  if (!sh || sh.getLastRow() < 2) return [];
  const v = sh.getRange(2, 3, sh.getLastRow() - 1, 4).getValues(); // 課題ID | 課題名 | 時間外 | ID
  const ids = [];
  for (let i = 0; i < v.length; i++) {
    if (String(v[i][3]).trim() === String(id).trim() && v[i][0] && ids.indexOf(String(v[i][0])) < 0) {
      ids.push(String(v[i][0]));
    }
  }
  return ids;
}

function allTasks_() {
  return taskSheetValues_().map(function (r) {
    const t = toTask_(r);
    return { id: t.id, name: t.name, subject: t.subject, unit: t.unit, file: t.file,
      grade: t.grade, cls: t.cls,
      open: t.open ? Utilities.formatDate(t.open, TZ, "yyyy-MM-dd HH:mm") : "",
      due:  t.due  ? Utilities.formatDate(t.due,  TZ, "yyyy-MM-dd HH:mm") : "",
      count: t.count, order: t.order, retry: t.retry, owner: t.owner, active: t.active };
  });
}

/* ====== 問題CSV（Driveフォルダ） ====== */
function csvFiles_() {
  if (!QUESTION_FOLDER_ID) return [];
  const it = DriveApp.getFolderById(QUESTION_FOLDER_ID).getFiles();
  const out = [];
  while (it.hasNext()) {
    const f = it.next();
    if (/\.csv$/i.test(f.getName())) out.push(f.getName());
  }
  return out.sort();
}

function questionsFor_(taskId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const t = findTask_(ss, taskId);
  if (!t) return { ok: false, msg: "課題が見つかりません。" };
  if (!QUESTION_FOLDER_ID) return { ok: false, msg: "問題フォルダが設定されていません。" };
  const it = DriveApp.getFolderById(QUESTION_FOLDER_ID).getFilesByName(t.file);
  if (!it.hasNext()) return { ok: false, msg: "問題ファイルが見つかりません：" + t.file };
  return { ok: true, csv: readCsv_(it.next()),
    task: { id: t.id, name: t.name, subject: t.subject, unit: t.unit,
            count: t.count, order: t.order } };
}

/** ExcelのShift-JIS保存にも対応する */
function readCsv_(file) {
  const blob = file.getBlob();
  let t = blob.getDataAsString("UTF-8");
  if (t.indexOf("\uFFFD") >= 0) {
    try { t = blob.getDataAsString("Shift_JIS"); } catch (e) {}
  }
  return t.replace(/^\uFEFF/, "");
}

/* ====== ログイン券 ====== */
function todayStr_() { return Utilities.formatDate(new Date(), TZ, "yyyyMMdd"); }
function yesterdayStr_() { return Utilities.formatDate(new Date(Date.now() - 86400000), TZ, "yyyyMMdd"); }
function ticketFor_(role, id, day) { return sha256_(SECRET + "|" + role + "|" + id + "|" + day); }
function verifyTicket_(role, id, ticket) {
  if (!id || !ticket) return false;
  const t = String(ticket).toLowerCase(), k = String(id).trim();
  return t === ticketFor_(role, k, todayStr_()) || t === ticketFor_(role, k, yesterdayStr_());
}

/* ====== 名簿・教員 ====== */
function lookupRoster_(ss, id) {
  const sh = ss.getSheetByName(SH_ROSTER);
  const miss = { found: false, grade: "", cls: "", no: "", pw: "" };
  if (!sh || sh.getLastRow() < 2) return miss;
  const v = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();
  const k = String(id).trim();
  for (let i = 0; i < v.length; i++) {
    if (String(v[i][0]).trim() === k) {
      return { found: true, grade: v[i][1], cls: v[i][2], no: v[i][3], pw: v[i][4] };
    }
  }
  return miss;
}
function lookupTeacher_(ss, id) {
  const sh = ss.getSheetByName(SH_TEACHER);
  const miss = { found: false, name: "", pw: "" };
  if (!sh || sh.getLastRow() < 2) return miss;
  const v = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
  const k = String(id).trim();
  for (let i = 0; i < v.length; i++) {
    if (String(v[i][0]).trim() === k) return { found: true, name: v[i][1], pw: v[i][2] };
  }
  return miss;
}

/* ====== ユーティリティ ====== */
function isHash_(v) { return /^[0-9a-fA-F]{64}$/.test(String(v).trim()); }
function sha256_(str) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, str, Utilities.Charset.UTF_8)
    .map(function (b) { return ("0" + (b & 0xFF).toString(16)).slice(-2); }).join("");
}
function textOut_(msg) {
  return ContentService.createTextOutput(msg).setMimeType(ContentService.MimeType.TEXT);
}
