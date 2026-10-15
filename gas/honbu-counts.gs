/* ============================================================
   本部ダッシュボード 件数API（2026-10）
   ・本部ダッシュボードのカードに出す「未処理の件数」だけを返す。
     中身（氏名・本文など）は一切返さない。
   ・置き場所：cryk0289 の Apps Script プロジェクト「本部ダッシュボード 件数API」
   ・デプロイ：種類「ウェブアプリ」／次のユーザーとして実行：自分（cryk0289）／アクセス：全員
     URLを変えずに更新するときは［デプロイ］→［デプロイを管理］→ 鉛筆 →
     バージョン「新バージョン」→［デプロイ］
   ・cryk0289 が読めるシートしか数えられない。ご意見箱（hojyokinsenyo 所有）は
     cryk0289 に閲覧以上の権限で共有しておくこと。
   ・2026-10-15：ご意見箱は職員ダッシュボードから「未対応」で書き込まれるため、
     「未対応」も未処理として数えるよう修正
   ============================================================ */

var CONFIG = {
  // 採用計画承認：採用申請システム（cryk0289）。①採用前申請と②応募者決裁の両方を数える
  saiyo: {
    spreadsheetId: '1mNWDy1eQZUEbnJc0R8roaJ3SXcvnw9Ej4xbxXvDjQ4w',
    sheets: [
      { name: '①採用前申請', statusHeader: 'ステータス' },
      { name: '②応募者決裁', statusHeader: 'ステータス' }
    ],
    // 「本部確認中」「決裁待ち」「未確認」など、中・待・未 を含む状態と、空欄を未処理として数える
    isPending: function (s) { return s === '' || /中|待|未/.test(s); }
  },
  // ご意見箱：シート名（タブ）は「目安箱」のまま
  goiken: {
    spreadsheetId: '1z3wCamSk7Mjm4hFrP_4D5Q9WKmXOJtUmXuFs0GEdeQI',
    sheets: [ { name: '目安箱', statusHeader: '状態' } ],
    // 「未対応」「受領」（届いただけで未対応）と空欄を未対応として数える。「確認」「対応」は数えない
    isPending: function (s) { return s === '' || s === '未対応' || s === '受領'; }
  },
  // 決裁承認：本部ダッシュボードは決裁承認GAS（「決裁申請 管理」）から直接件数を取るため、ここでは数えない
  kessai: null
};

var CACHE_SECONDS = 60;   // 同じ結果を60秒使い回す（シートを毎回開かないように）

function doGet(e) {
  var data;
  var cache = CacheService.getScriptCache();
  var hit = cache.get('counts');
  if (hit) {
    data = JSON.parse(hit);
  } else {
    data = { ok: true, updatedAt: new Date().toISOString(), counts: {} };
    Object.keys(CONFIG).forEach(function (key) {
      data.counts[key] = CONFIG[key] ? countPending_(CONFIG[key]) : null;   // null＝未連携
    });
    cache.put('counts', JSON.stringify(data), CACHE_SECONDS);
  }
  var json = JSON.stringify(data);
  var cb = (e && e.parameter && e.parameter.callback) || '';
  if (/^[A-Za-z_$][\w$]{0,40}$/.test(cb)) {
    return ContentService.createTextOutput(cb + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

/* 1つの項目の未処理件数を数える。読めなかったときは -1（画面では「取得エラー」） */
function countPending_(cfg) {
  try {
    var ss = SpreadsheetApp.openById(cfg.spreadsheetId);
    var total = 0;
    cfg.sheets.forEach(function (sc) {
      var sh = ss.getSheetByName(sc.name);
      if (!sh) throw new Error('シートが見つかりません：' + sc.name);
      var values = sh.getDataRange().getDisplayValues();
      if (values.length < 2) return;
      var col = values[0].indexOf(sc.statusHeader);
      if (col < 0) throw new Error('列が見つかりません：' + sc.statusHeader);
      for (var r = 1; r < values.length; r++) {
        if (String(values[r][0]).trim() === '') continue;   // A列が空の行（空行）は数えない
        if (cfg.isPending(String(values[r][col]).trim())) total++;
      }
    });
    return total;
  } catch (err) {
    console.error(err);
    return -1;
  }
}

/* エディタから実行して動作確認する用（実行ログに件数が出る） */
function testCounts() {
  CacheService.getScriptCache().remove('counts');
  Logger.log(doGet({ parameter: {} }).getContent());
}
