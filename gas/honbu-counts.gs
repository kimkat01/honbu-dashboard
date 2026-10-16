/* ============================================================
   本部ダッシュボード 件数API ＋ 2段階認証（2026-10）
   ・置き場所：cryk0289 の Apps Script プロジェクト「本部ダッシュボード 件数API」
     （このファイルの中身を、既存のコードと全部入れ替える）
   ・デプロイ：種類「ウェブアプリ」／次のユーザーとして実行：自分（cryk0289）／アクセス：全員
     URLを変えずに更新するときは［デプロイ］→［デプロイを管理］→ 鉛筆 →
     バージョン「新バージョン」→［デプロイ］

   【2段階認証のしくみ】
   ① ログインID＋パスワード → 正しければ、登録メールに6桁の確認コードを送る
   ② 確認コード → 正しければ「通行証（トークン）」を発行（12時間有効）
   ③ 件数の取得は、正しい通行証があるときだけ返す
   ・パスワードはこのコードにもシートにも平文で残さない（スクリプトプロパティに
     ハッシュ値だけを保存）
   ・パスワードを5回まちがえると、そのIDは15分間ロック
   ・確認コードは10分間有効・5回まちがえると無効

   【最初にやること】
   1. 下の HONBU_USERS に3人のメールアドレスを入れて保存
   2. 関数 initPasswords を選んで［実行］→ 権限を許可
      → 実行ログに3人の初期パスワードが1回だけ表示されるので、控えて本人に渡す
      （ログは後から見返さない前提。紙などで渡したら控えは処分）
   3. 上の手順で「新バージョン」としてデプロイ
   ・パスワードを忘れた人がいたら：resetPasswordHonbu1 などを実行すると、
     その人だけ新しい仮パスワードが実行ログに出る

   【件数APIの設定（従来どおり）】
   ・cryk0289 が読めるシートしか数えられない。ご意見箱（hojyokinsenyo 所有）は
     cryk0289 に閲覧以上の権限で共有しておくこと。
   ・2026-10-15：ご意見箱は職員ダッシュボードから「未対応」で書き込まれるため、
     「未対応」も未処理として数えるよう修正
   ============================================================ */

/* ---------- 本部の3人（ここだけ書き換える） ---------- */
var HONBU_USERS = [
  { id: 'honbu1', name: '理事長',     email: '' },   // ← 理事長のメールアドレス
  { id: 'honbu2', name: '理事長補佐', email: '' },   // ← 理事長補佐のメールアドレス
  { id: 'honbu3', name: '担当',       email: '' }    // ← 担当のメールアドレス
];

var AUTH = {
  CODE_MINUTES: 10,     // 確認コードの有効時間（分）
  CODE_TRIES: 5,        // 確認コードの入力ミス許容回数
  LOCK_FAILS: 5,        // パスワードを何回まちがえたらロックするか
  LOCK_MINUTES: 15,     // ロック時間（分）
  SESSION_HOURS: 12,    // 通行証（ログイン状態）の有効時間
  HASH_ROUNDS: 2000     // パスワードのハッシュ化の繰り返し回数
};

/* ---------- 件数APIの設定（従来どおり） ---------- */
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
var AUTH_VERSION = 1;     // 画面側が「2段階認証が有効なGASか」を見分ける目印

/* ============================================================
   入口
   ============================================================ */

/* GET：件数の取得（通行証が必要）／通行証の確認 */
function doGet(e) {
  var p = (e && e.parameter) || {};
  var res;
  if (p.action === 'verify') {
    res = verifyToken_(p.token);
  } else {
    var who = checkToken_(p.token);
    res = who ? getCounts_() : { ok: false, error: 'auth', authVersion: AUTH_VERSION };
  }
  return output_(res, p.callback);
}

/* POST：ログイン（パスワード → 確認コード）／ログアウト
   画面からは Content-Type: text/plain で JSON を送る（パスワードをURLに載せない） */
function doPost(e) {
  var body = {};
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) {}
  var lock = LockService.getScriptLock();
  var res;
  try {
    lock.waitLock(10000);
    if (body.action === 'login')       res = login_(String(body.id || ''), String(body.password || ''));
    else if (body.action === 'code')   res = verifyCode_(String(body.pending || ''), String(body.code || ''));
    else if (body.action === 'logout') res = logout_(String(body.token || ''));
    else res = { ok: false, error: 'bad_request' };
  } catch (err) {
    console.error(err);
    res = { ok: false, error: 'server' };
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
  res.authVersion = AUTH_VERSION;
  return output_(res, '');
}

function output_(obj, cb) {
  var json = JSON.stringify(obj);
  if (cb && /^[A-Za-z_$][\w$]{0,40}$/.test(cb)) {
    return ContentService.createTextOutput(cb + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

/* ============================================================
   2段階認証
   ============================================================ */

function props_() { return PropertiesService.getScriptProperties(); }

function findUser_(id) {
  id = String(id || '').trim().toLowerCase();
  for (var i = 0; i < HONBU_USERS.length; i++) {
    if (HONBU_USERS[i].id.toLowerCase() === id) return HONBU_USERS[i];
  }
  return null;
}

/* ① ID＋パスワード → 確認コードをメール送信 */
function login_(id, password) {
  cleanup_();
  var user = findUser_(id);
  var P = props_();
  var now = Date.now();

  // ロック中か（存在しないIDも同じ扱いにして、IDの有無を外から分からなくする）
  var lockKey = 'lock_' + (user ? user.id : 'x_' + id.toLowerCase().slice(0, 20));
  var lockInfo = readJson_(P.getProperty(lockKey)) || { fails: 0, until: 0 };
  if (lockInfo.until > now) {
    return { ok: false, error: 'locked', minutes: Math.ceil((lockInfo.until - now) / 60000) };
  }

  var stored = user ? readJson_(P.getProperty('pw_' + user.id)) : null;
  var good = !!(user && stored && hash_(password, stored.salt) === stored.hash);
  if (!good) {
    lockInfo.fails = (lockInfo.fails || 0) + 1;
    if (lockInfo.fails >= AUTH.LOCK_FAILS) {
      lockInfo = { fails: 0, until: now + AUTH.LOCK_MINUTES * 60000 };
      P.setProperty(lockKey, JSON.stringify(lockInfo));
      return { ok: false, error: 'locked', minutes: AUTH.LOCK_MINUTES };
    }
    P.setProperty(lockKey, JSON.stringify(lockInfo));
    return { ok: false, error: 'wrong', left: AUTH.LOCK_FAILS - lockInfo.fails };
  }
  P.deleteProperty(lockKey);

  if (!user.email) return { ok: false, error: 'no_email' };

  // 確認コードを作って送る
  var code = randomDigits_(6);
  var pending = randomToken_();
  P.setProperty('otp_' + pending, JSON.stringify({
    id: user.id,
    hash: quick_(pending + ':' + code),
    exp: now + AUTH.CODE_MINUTES * 60000,
    tries: 0
  }));

  MailApp.sendEmail({
    to: user.email,
    subject: '【本部ダッシュボード】確認コード ' + code,
    name: 'しおどめ 本部ダッシュボード',
    body:
      user.name + ' 様\n\n' +
      '本部ダッシュボードの確認コードは次のとおりです。\n\n' +
      '　　' + code + '\n\n' +
      'このコードの有効時間は' + AUTH.CODE_MINUTES + '分です。\n' +
      '心当たりのない場合は、パスワードが他の人に知られている可能性があります。\n' +
      'このメールは破棄し、担当者にご連絡ください。\n'
  });

  return { ok: true, pending: pending, email: maskEmail_(user.email), minutes: AUTH.CODE_MINUTES };
}

/* ② 確認コード → 通行証を発行 */
function verifyCode_(pending, code) {
  var P = props_();
  var key = 'otp_' + pending;
  var otp = readJson_(P.getProperty(key));
  if (!pending || !otp) return { ok: false, error: 'expired' };
  if (otp.exp < Date.now()) { P.deleteProperty(key); return { ok: false, error: 'expired' }; }

  code = String(code || '').replace(/\D/g, '');
  if (quick_(pending + ':' + code) !== otp.hash) {
    otp.tries++;
    if (otp.tries >= AUTH.CODE_TRIES) { P.deleteProperty(key); return { ok: false, error: 'expired' }; }
    P.setProperty(key, JSON.stringify(otp));
    return { ok: false, error: 'wrong_code', left: AUTH.CODE_TRIES - otp.tries };
  }
  P.deleteProperty(key);

  var user = findUser_(otp.id);
  if (!user) return { ok: false, error: 'expired' };
  var token = randomToken_() + randomToken_();
  var exp = Date.now() + AUTH.SESSION_HOURS * 3600000;
  P.setProperty('tok_' + quick_(token), JSON.stringify({ id: user.id, exp: exp }));
  console.log('本部ダッシュボード ログイン：' + user.name + '（' + user.id + '）');
  return { ok: true, token: token, name: user.name, exp: exp };
}

/* 通行証の確認（正しければ利用者を返す。だめなら null） */
function checkToken_(token) {
  if (!token || String(token).length < 40) return null;
  var key = 'tok_' + quick_(String(token));
  var t = readJson_(props_().getProperty(key));
  if (!t) return null;
  if (t.exp < Date.now()) { props_().deleteProperty(key); return null; }
  var user = findUser_(t.id);
  return user ? { user: user, exp: t.exp } : null;
}

function verifyToken_(token) {
  var who = checkToken_(token);
  if (!who) return { ok: false, error: 'auth', authVersion: AUTH_VERSION };
  return { ok: true, name: who.user.name, exp: who.exp, authVersion: AUTH_VERSION };
}

function logout_(token) {
  if (token) props_().deleteProperty('tok_' + quick_(token));
  return { ok: true };
}

/* 期限切れの確認コード・通行証・ロックを片付ける */
function cleanup_() {
  var P = props_();
  var all = P.getProperties();
  var now = Date.now();
  Object.keys(all).forEach(function (k) {
    if (!/^(otp_|tok_|lock_)/.test(k)) return;
    var v = readJson_(all[k]);
    if (!v) { P.deleteProperty(k); return; }
    if (k.indexOf('lock_') === 0) { if (v.until && v.until < now && !v.fails) P.deleteProperty(k); return; }
    if (v.exp && v.exp < now) P.deleteProperty(k);
  });
}

/* ---------- 小道具 ---------- */
function hash_(text, salt) {
  var base = String(salt) + ':' + String(text);
  var h = base;
  for (var i = 0; i < AUTH.HASH_ROUNDS; i++) {
    h = Utilities.base64EncodeWebSafe(
      Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h + '|' + base, Utilities.Charset.UTF_8));
  }
  return h;
}

/* 通行証・確認コード用の1回だけのハッシュ（推測できない長い値なので繰り返し不要） */
function quick_(text) {
  return Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(text), Utilities.Charset.UTF_8));
}

function randomToken_() {
  return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
}

function randomDigits_(n) {
  var hex = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  var out = '';
  for (var i = 0; out.length < n && i < hex.length; i++) {
    var d = parseInt(hex.charAt(i), 16);
    if (d < 10) out += d;            // 0〜9 だけ使う（偏りを出さない）
  }
  while (out.length < n) out += String(Math.floor(Math.random() * 10));
  return out;
}

function randomPassword_() {
  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789';   // 0/O・1/l/I など紛らわしい字は除く
  var hex = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  var out = '';
  for (var i = 0; i + 1 < hex.length && out.length < 10; i += 2) {
    out += chars.charAt(parseInt(hex.substr(i, 2), 16) % chars.length);
  }
  return out;
}

function maskEmail_(email) {
  var m = String(email).split('@');
  if (m.length !== 2) return '登録メール';
  return m[0].charAt(0) + '***@' + m[1];
}

function readJson_(s) { try { return s ? JSON.parse(s) : null; } catch (e) { return null; } }

function setPassword_(user, password) {
  var salt = randomToken_();
  props_().setProperty('pw_' + user.id, JSON.stringify({ salt: salt, hash: hash_(password, salt), at: new Date().toISOString() }));
  props_().deleteProperty('lock_' + user.id);
}

/* ============================================================
   エディタから実行する関数
   ============================================================ */

/* 初回だけ：まだパスワードが無い人に初期パスワードを作り、実行ログに表示する */
function initPasswords() {
  HONBU_USERS.forEach(function (u) {
    if (props_().getProperty('pw_' + u.id)) {
      Logger.log(u.name + '（ID：' + u.id + '）はパスワード設定済みのため変更しません');
      return;
    }
    var pw = randomPassword_();
    setPassword_(u, pw);
    Logger.log(u.name + '　ログインID：' + u.id + '　初期パスワード：' + pw + (u.email ? '' : '　※メールアドレス未登録'));
  });
}

/* パスワードを忘れた人の再発行（その人だけ新しい仮パスワードが実行ログに出る） */
function resetPasswordHonbu1() { resetPassword_('honbu1'); }
function resetPasswordHonbu2() { resetPassword_('honbu2'); }
function resetPasswordHonbu3() { resetPassword_('honbu3'); }
function resetPassword_(id) {
  var u = findUser_(id);
  if (!u) { Logger.log('IDが見つかりません：' + id); return; }
  var pw = randomPassword_();
  setPassword_(u, pw);
  logoutAll_(u.id);
  Logger.log(u.name + '　ログインID：' + u.id + '　新しい仮パスワード：' + pw);
}

/* 全員を強制ログアウト（スマホ・PCを紛失したときなど） */
function logoutEveryone() { logoutAll_(null); Logger.log('全員をログアウトしました'); }
function logoutAll_(id) {
  var P = props_();
  var all = P.getProperties();
  Object.keys(all).forEach(function (k) {
    if (k.indexOf('tok_') !== 0) return;
    var v = readJson_(all[k]);
    if (!id || (v && v.id === id)) P.deleteProperty(k);
  });
}

/* メールが届くかの確認用（担当＝honbu3 にテストメールを送る） */
function testMail() {
  var u = findUser_('honbu3');
  if (!u || !u.email) { Logger.log('honbu3 のメールアドレスが未登録です'); return; }
  MailApp.sendEmail(u.email, '【本部ダッシュボード】テストメール', '確認コードのメールはこの差出人から届きます。');
  Logger.log('送信しました：' + u.email);
}

/* ============================================================
   件数API（従来どおり）
   ============================================================ */

function getCounts_() {
  var cache = CacheService.getScriptCache();
  var hit = cache.get('counts');
  if (hit) return JSON.parse(hit);
  var data = { ok: true, updatedAt: new Date().toISOString(), counts: {} };
  Object.keys(CONFIG).forEach(function (key) {
    data.counts[key] = CONFIG[key] ? countPending_(CONFIG[key]) : null;   // null＝未連携
  });
  cache.put('counts', JSON.stringify(data), CACHE_SECONDS);
  return data;
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

/* エディタから実行して動作確認する用（実行ログに件数が出る。認証なしで中身を確認） */
function testCounts() {
  CacheService.getScriptCache().remove('counts');
  Logger.log(JSON.stringify(getCounts_()));
}
