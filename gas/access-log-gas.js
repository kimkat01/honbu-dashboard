/* ============================================================
   しおどめ 利用状況（アクセスログ）GAS
   ------------------------------------------------------------
   ・置き場所：スプレッドシート「アクセスログ」（cryk0289）に紐付いた Apps Script
   ・役割：
     doPost … 園長ダッシュボード／職員ダッシュボードでツールを開いたときの記録を1行追加する
     doGet  … 本部ダッシュボードの「利用状況」画面（riyo-jokyo.html）に、記録と職員名簿を渡す
   ・職員名簿は「PINマスタ_テンプレート」のログインマスタから
     A列＝園名／B列＝ログインID／C列＝名前／G列＝権限 だけを読む（パスワード列は読まない）
   ・最初に1回だけ setup を実行してください（シートの見出し作成・権限の承認・月1回の整理の予約）
   ============================================================ */

var LOGIN_MASTER_ID = '1KYQJl14dvvWPM02xMzNye5hFFWiFU_uus-MVH3E-LNY'; // PINマスタ_テンプレート
var LOGIN_SHEET     = 'ログインマスタ';
var VIEW_KEY        = 'PgAzmLVXiEdDE3rDNSLjeY6J';   // 利用状況画面だけが知っている合言葉
var LOG_SHEET       = 'ログ';
var ARCHIVE_SHEET   = 'ログ_過去分';
var KEEP_DAYS       = 180;   // これより古い記録は月1回「ログ_過去分」へ移す
var MAX_VIEW_DAYS   = 90;    // 画面に渡す最大日数
var DASHES          = ['職員', '園長'];
var HEADER          = ['日時', 'ダッシュボード', '園名', 'ログインID', '名前', '権限', 'ツール'];

/* ---------- 最初に1回だけ実行 ---------- */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone('Asia/Tokyo');
  prepareSheet_(ss, LOG_SHEET);
  prepareSheet_(ss, ARCHIVE_SHEET);
  // 名簿が読めるか確認（ここで権限の承認画面が出ます）
  SpreadsheetApp.openById(LOGIN_MASTER_ID).getSheetByName(LOGIN_SHEET).getLastRow();
  // 月1回の整理（古い記録の移動）を予約。二重に予約しないよう、先に消してから作る
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'archiveOld') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('archiveOld').timeBased().onMonthDay(1).atHour(3).create();
}

function prepareSheet_(ss, name) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADER);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADER.length).setFontWeight('bold');
    sh.getRange('A:A').setNumberFormat('yyyy/mm/dd hh:mm:ss');
    sh.getRange('D:D').setNumberFormat('@');
  }
  return sh;
}

/* ---------- 記録を受け取る ---------- */
function doPost(e) {
  try {
    var d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (d.action !== 'log') return text_('ng');
    var dash = String(d.dash || '');
    var tool = clip_(d.tool, 40);
    var loginId = clip_(d.loginId, 20);
    if (DASHES.indexOf(dash) < 0 || !tool || !loginId) return text_('ng');
    var row = [new Date(), dash, clip_(d.garden, 40), loginId, clip_(d.name, 40), clip_(d.role, 20), tool];
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

/* ---------- 利用状況画面へ渡す ---------- */
function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.key !== VIEW_KEY) return json_({ ok: false, error: 'key' });
  var days = Math.min(MAX_VIEW_DAYS, Math.max(1, parseInt(p.days, 10) || 30));
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
      events.push([t, String(r[1]), String(r[2]), String(r[3]), String(r[4]), String(r[5]), String(r[6])]);
    }
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
  } catch (err) {}

  return json_({ ok: true, now: Date.now(), days: days, events: events, roster: roster });
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
