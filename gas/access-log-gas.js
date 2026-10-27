/* ============================================================
   しおどめ 利用状況（アクセスログ）GAS
   ------------------------------------------------------------
   ・置き場所：スプレッドシート「アクセスログ」（cryk0289）に紐付いた Apps Script
   ・役割：
     doPost … 園長ダッシュボード／職員ダッシュボードでツールを開いたときの記録を1行追加する
     doGet  … 本部ダッシュボードの「利用状況」画面（riyo-jokyo.html）に、記録と職員名簿を渡す
   ・職員名簿は「PINマスタ_テンプレート」のログインマスタから
     A列＝園名／B列＝ログインID／C列＝名前／G列＝権限 だけを読む（パスワード列は読まない）
   ・最初に1回だけ setup を実行してください（シートの見出し作成・権限の承認・予約の登録）
   ・2026-10-16：表示の高速化
     利用状況画面に渡す内容を「作り置き」（CacheService）するようにした。
     5分ごとの予約（warmCache）で作り置きを更新するので、画面を開いたときはそれを渡すだけになる。
     画面の「最新にする」（fresh=1）のときだけ、その場でシートを読み直す。
     職員名簿は変更が少ないので6時間使い回す。
     ※ このコードに貼り替えたら、setup をもう一度実行して予約を登録し直すこと
   ・2026-10-21：H列「操作」を追加（閲覧／投稿）。
     ツールを開いたときは「閲覧」、ご意見箱の送信・決裁申請・チャット送信などは「投稿」と記録する。
     それより前の行（H列が空）は「閲覧」として扱う。
     ※ 貼り替えたら setup を1回実行（ログのH1に「操作」の見出しが入ります）→ デプロイを「新バージョン」で更新
   ・2026-10-27：セキュリティ対策（利用状況画面の合言葉を廃止）
     これまで利用状況画面（riyo-jokyo.html）のソースに合言葉（VIEW_KEY）が書かれていて、
     URLを知っている人なら誰でも全職員の名簿と閲覧記録を取り出せた。
     → 画面で本部アカウント（ログインマスタのG列が「本部」の行）のログインIDとパスワードを入れ、
       このGASが照合して「通行証」を発行する方式に変更。記録と名簿は通行証があるときだけ渡す。
     ・通行証は30日有効（そのパソコンのブラウザに保存）。スクリプトプロパティに「vt_…」として保存する
     ・パスワードを5回続けて間違えると、そのログインIDは15分間ログインできない
     ・記録（action=log）の受け付けは今までどおり（園長・職員ダッシュボードは変更不要）
     ・doGet（合言葉で読む古い入口）は何も返さない
     ※ 貼り替えたら setup は不要。デプロイを「新バージョン」で更新するだけでよい
   ============================================================ */

var LOGIN_MASTER_ID = '1KYQJl14dvvWPM02xMzNye5hFFWiFU_uus-MVH3E-LNY'; // PINマスタ_テンプレート
var LOGIN_SHEET     = 'ログインマスタ';
var HONBU_ROLE      = '本部';      // ログインマスタG列がこの値の行だけ、利用状況画面に入れる
var TOKEN_DAYS      = 30;          // 通行証の有効日数
var FAIL_LIMIT      = 5;           // パスワードをこの回数続けて間違えたら一時停止
var FAIL_LOCK_SEC   = 900;         // 一時停止の時間（15分）
var LOG_SHEET       = 'ログ';
var ARCHIVE_SHEET   = 'ログ_過去分';
var KEEP_DAYS       = 180;   // これより古い記録は月1回「ログ_過去分」へ移す
var MAX_VIEW_DAYS   = 90;    // 画面に渡す最大日数
var DASHES          = ['職員', '園長'];
var HEADER          = ['日時', 'ダッシュボード', '園名', 'ログインID', '名前', '権限', 'ツール', '操作'];
var ACTS            = ['閲覧', '投稿'];

var VIEW_CACHE_SEC   = 600;     // 作り置きの有効時間（予約で5分ごとに作り直す）
var ROSTER_CACHE_SEC = 21600;   // 職員名簿の使い回し時間（6時間）
var CHUNK            = 90000;   // CacheService は1件100KBまでなので、分けて入れる

/* ---------- 最初に1回だけ実行（貼り替えたときももう一度実行） ---------- */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone('Asia/Tokyo');
  prepareSheet_(ss, LOG_SHEET);
  prepareSheet_(ss, ARCHIVE_SHEET);
  // 名簿が読めるか確認（ここで権限の承認画面が出ます）
  SpreadsheetApp.openById(LOGIN_MASTER_ID).getSheetByName(LOGIN_SHEET).getLastRow();
  // 予約：月1回の整理（古い記録の移動）と、5分ごとの作り置き更新。二重に予約しないよう、先に消してから作る
  ScriptApp.getProjectTriggers().forEach(function (t) {
    var f = t.getHandlerFunction();
    if (f === 'archiveOld' || f === 'warmCache') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('archiveOld').timeBased().onMonthDay(1).atHour(3).create();
  ScriptApp.newTrigger('warmCache').timeBased().everyMinutes(5).create();
  warmCache();
}

function prepareSheet_(ss, name) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADER);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADER.length).setFontWeight('bold');
    sh.getRange('A:A').setNumberFormat('yyyy/mm/dd hh:mm:ss');
    sh.getRange('D:D').setNumberFormat('@');
  } else if (sh.getRange(1, HEADER.length).getValue() === '') {
    // 2026-10-21：前からあるシートに「操作」の見出しを足す
    sh.getRange(1, HEADER.length).setValue(HEADER[HEADER.length - 1]).setFontWeight('bold');
  }
  return sh;
}

/* ---------- 記録を受け取る ---------- */
function doPost(e) {
  var d;
  try { d = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) { return text_('ng'); }
  // 利用状況画面（本部）からの入口
  if (d.action === 'login')  return json_(login_(d.loginId, d.password));
  if (d.action === 'view')   return viewOut_(d);
  if (d.action === 'logout') { if (d.token) props_().deleteProperty(tokKey_(d.token)); return json_({ ok: true }); }
  return logPost_(d);
}

function logPost_(d) {
  try {
    if (d.action !== 'log') return text_('ng');
    var dash = String(d.dash || '');
    var tool = clip_(d.tool, 40);
    var loginId = clip_(d.loginId, 20);
    if (DASHES.indexOf(dash) < 0 || !tool || !loginId) return text_('ng');
    var act = ACTS.indexOf(String(d.act || '')) >= 0 ? String(d.act) : '閲覧';
    var row = [new Date(), dash, clip_(d.garden, 40), loginId, clip_(d.name, 40), clip_(d.role, 20), tool, act];
    var lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      prepareSheet_(ss, LOG_SHEET).appendRow(row);
    } finally {
      lock.releaseLock();
    }
    return text_('ok');
  } catch (err) {
    return text_('error');
  }
}

/* ---------- 古い入口（合言葉で読む）：2026-10-27 廃止 ---------- */
function doGet(e) {
  return json_({ ok: false, error: 'auth' });
}

/* ---------- 利用状況画面へ渡す（通行証が必要） ---------- */
function viewOut_(d) {
  if (!checkToken_(d.token)) return json_({ ok: false, error: 'auth' });
  var days = Math.min(MAX_VIEW_DAYS, Math.max(1, parseInt(d.days, 10) || 30));
  var fresh = String(d.fresh) === '1' || d.fresh === true;
  if (!fresh) {
    var hit = cacheGet_('view_' + days);
    if (hit) return ContentService.createTextOutput(hit).setMimeType(ContentService.MimeType.JSON);
  }
  var json = JSON.stringify(buildPayload_(days, fresh));
  cachePut_('view_' + days, json, VIEW_CACHE_SEC);
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- 本部アカウントのログイン（2026-10-27） ---------- */
function login_(loginId, password) {
  var id = norm_(loginId), pw = String(password == null ? '' : password).normalize('NFKC').trim();
  if (!/^[0-9]{4}$/.test(id) || !pw) return { ok: false, error: 'input', message: 'ログインIDとパスワードを入力してください' };

  var cache = CacheService.getScriptCache();
  var failKey = 'fail_' + id;
  var fails = parseInt(cache.get(failKey), 10) || 0;
  if (fails >= FAIL_LIMIT) return { ok: false, error: 'locked', message: 'パスワードを続けて間違えたため、15分間ログインできません' };

  var sh = SpreadsheetApp.openById(LOGIN_MASTER_ID).getSheetByName(LOGIN_SHEET);
  var vals = sh.getRange(2, 1, Math.max(1, sh.getLastRow() - 1), 7).getValues();
  var hit = null;
  for (var i = 0; i < vals.length; i++) {
    var r = vals[i];
    if (norm_(r[1]) !== id || norm_(r[6]) !== HONBU_ROLE) continue;
    var stored = String(r[3] == null ? '' : r[3]).trim();
    if (stored && stored === pw) { hit = r; break; }
  }
  if (!hit) {
    cache.put(failKey, String(fails + 1), FAIL_LOCK_SEC);
    return { ok: false, error: 'bad', message: 'ログインIDまたはパスワードが正しくありません（本部アカウントだけが入れます）' };
  }
  cache.remove(failKey);
  pruneTokens_();
  var token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  var exp = Date.now() + TOKEN_DAYS * 86400000;
  props_().setProperty(tokKey_(token), JSON.stringify({ id: id, name: String(hit[2]).trim(), exp: exp }));
  return { ok: true, token: token, name: String(hit[2]).trim(), exp: exp };
}

function checkToken_(token) {
  if (!token || String(token).length < 40) return null;
  var key = tokKey_(token);
  var raw = props_().getProperty(key);
  if (!raw) return null;
  var t = null;
  try { t = JSON.parse(raw); } catch (e) {}
  if (!t || !t.exp || t.exp < Date.now()) { props_().deleteProperty(key); return null; }
  return t;
}

/* 期限切れの通行証を片付ける */
function pruneTokens_() {
  try {
    var p = props_(), all = p.getProperties(), now = Date.now();
    Object.keys(all).forEach(function (k) {
      if (k.indexOf('vt_') !== 0) return;
      var t = null;
      try { t = JSON.parse(all[k]); } catch (e) {}
      if (!t || !t.exp || t.exp < now) p.deleteProperty(k);
    });
  } catch (e) {}
}

function props_() { return PropertiesService.getScriptProperties(); }
function tokKey_(token) {
  var b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(token), Utilities.Charset.UTF_8);
  return 'vt_' + b.map(function (x) { return ('0' + (x & 255).toString(16)).slice(-2); }).join('');
}
function norm_(s) { return String(s == null ? '' : s).normalize('NFKC').trim(); }

/* 画面に渡す中身を作る（記録＋職員名簿） */
function buildPayload_(days, freshRoster) {
  var since = Date.now() - days * 86400000;
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET);
  var events = [];
  if (sh && sh.getLastRow() > 1) {
    var vals = sh.getRange(2, 1, sh.getLastRow() - 1, HEADER.length).getValues();
    for (var i = vals.length - 1; i >= 0; i--) {   // 新しい順に読み、期間外に出たら止める
      var r = vals[i];
      var t = r[0] instanceof Date ? r[0].getTime() : new Date(r[0]).getTime();
      if (!t) continue;
      if (t < since) break;
      events.push([t, String(r[1]), String(r[2]), String(r[3]), String(r[4]), String(r[5]), String(r[6]), String(r[7] || '閲覧')]);
    }
  }
  return { ok: true, now: Date.now(), days: days, events: events, roster: getRoster_(freshRoster) };
}

/* 職員名簿（6時間使い回す） */
function getRoster_(fresh) {
  if (!fresh) {
    var hit = cacheGet_('roster');
    if (hit) { try { return JSON.parse(hit); } catch (e) {} }
  }
  var roster = [];
  try {
    var ms = SpreadsheetApp.openById(LOGIN_MASTER_ID).getSheetByName(LOGIN_SHEET);
    if (ms && ms.getLastRow() > 1) {
      var mv = ms.getRange(2, 1, ms.getLastRow() - 1, 7).getValues();   // A〜G。D列（パスワード）は使わない
      mv.forEach(function (r) {
        var garden = String(r[0]).trim(), id = String(r[1]).trim(), name = String(r[2]).trim();
        if (garden && id && name) roster.push([garden, id, name, String(r[6]).trim()]);
      });
    }
    if (roster.length) cachePut_('roster', JSON.stringify(roster), ROSTER_CACHE_SEC);
  } catch (err) {}
  return roster;
}

/* ---------- 5分ごとの予約：作り置きを更新 ---------- */
function warmCache() {
  var json = JSON.stringify(buildPayload_(MAX_VIEW_DAYS, false));
  cachePut_('view_' + MAX_VIEW_DAYS, json, VIEW_CACHE_SEC);
}

/* ---------- 作り置き（100KBを超えるときは分けて入れる） ---------- */
function cachePut_(key, str, sec) {
  try {
    var cache = CacheService.getScriptCache();
    var n = Math.ceil(str.length / CHUNK) || 1;
    if (n > 20) return;   // 大きすぎるときは作り置きしない
    var obj = {};
    obj[key + '__n'] = String(n);
    for (var i = 0; i < n; i++) obj[key + '__' + i] = str.substr(i * CHUNK, CHUNK);
    cache.putAll(obj, sec);
  } catch (e) {}
}
function cacheGet_(key) {
  try {
    var cache = CacheService.getScriptCache();
    var n = parseInt(cache.get(key + '__n'), 10);
    if (!n) return null;
    var keys = [];
    for (var i = 0; i < n; i++) keys.push(key + '__' + i);
    var got = cache.getAll(keys), out = '';
    for (var j = 0; j < n; j++) { if (got[keys[j]] == null) return null; out += got[keys[j]]; }
    return out;
  } catch (e) { return null; }
}

/* ---------- 月1回：古い記録を「ログ_過去分」へ移す ---------- */
function archiveOld() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(LOG_SHEET);
  if (!sh || sh.getLastRow() < 2) return;
  var limit = Date.now() - KEEP_DAYS * 86400000;
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, HEADER.length).getValues();
  var n = 0;
  while (n < vals.length) {
    var t = vals[n][0] instanceof Date ? vals[n][0].getTime() : new Date(vals[n][0]).getTime();
    if (t && t >= limit) break;
    n++;
  }
  if (n === 0) return;
  var arc = prepareSheet_(ss, ARCHIVE_SHEET);
  arc.getRange(arc.getLastRow() + 1, 1, n, HEADER.length).setValues(vals.slice(0, n));
  sh.deleteRows(2, n);
}

/* ---------- 小道具 ---------- */
function clip_(v, n) { return String(v == null ? '' : v).replace(/[\r\n\t]/g, ' ').trim().slice(0, n); }
function text_(s) { return ContentService.createTextOutput(s).setMimeType(ContentService.MimeType.TEXT); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
