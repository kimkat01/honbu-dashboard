/**
 * 決裁承認システム（GAS） 2026-10
 * ------------------------------------------------------------
 * 置き場所：cryk0289 アカウントで新規作成したスプレッドシート
 *           「決裁申請 管理」 → 拡張機能 → Apps Script に貼り付け
 *
 * 流れ：園長が申請（kessai-shinsei.html）
 *        → 法人本部が「承認／差戻し／却下」（kessai-shonin.html）
 *        → 理事長が「承認／差戻し／却下」
 *        → 決裁済
 *   差戻し：園長にメールで修正用リンクを送り、修正して再申請 → 再び本部から
 *
 * 通知メール：①申請・再申請時に本部へ ②結果（決裁済・差戻し・却下）を園長へ
 *
 * 2026-10-16：合言葉を廃止。本部ダッシュボードの2段階認証でログインした人だけが
 *             承認画面を使える（通行証を「本部ダッシュボード 件数API」に問い合わせて確認）
 *             ※先に「本部ダッシュボード 件数API」を2段階認証版にしておくこと
 *
 * 初回だけ：下の「設定」を書き換え → setup() を実行して権限を許可
 *           → デプロイ → 新しいデプロイ → ウェブアプリ
 *             （実行ユーザー：自分 ／ アクセス：全員）
 * コードを変えたら：デプロイを管理 → 編集 → バージョン「新バージョン」
 *                   （URLは変わらない）
 * ------------------------------------------------------------
 */

/* ======================== 設定 ======================== */
const CONFIG = {
  // 本部ダッシュボードのログイン確認先（「本部ダッシュボード 件数API」の /exec URL）
  AUTH_URL: 'https://script.google.com/macros/s/AKfycbzJHhK-jajUwamz69kCWfZu6H2-L9Z4_Hos7O5VrQFd_CHVrMH1HT5AvhgzPIBU-Kx4/exec',

  // 本部ダッシュボードのログインIDごとに、押せる承認欄を決める
  //   'rijicho'＝「理事長」欄だけ押せる ／ 'honbu'＝「法人本部」欄だけ押せる
  //   ここに無いIDの人は承認画面に入れない
  ROLE_BY_ID: {
    honbu1: 'rijicho',   // 理事長
    honbu2: 'honbu',     // 理事長補佐
    honbu3: 'honbu'      // 担当
  },

  // 申請があったときに通知する本部のメールアドレス（複数はカンマ区切り）
  HONBU_EMAIL: 'cryk0289@gmail.com,akihiko.shiba4848@gmail.com',

  // 画面のURL（メール本文のリンクに使う）
  FORM_URL:    'https://kimkat01.github.io/shiodome-knowledge/kessai-shinsei.html',
  APPROVE_URL: 'https://kimkat01.github.io/honbu-dashboard/kessai-shonin.html',

  // 見積書などの添付ファイルを「リンクを知っている人は閲覧可」にするか
  //   true ：承認画面のリンクから誰でも開ける（理事長がGoogleにログインしていなくても見られる）
  //   false：自分（cryk0289）と、フォルダを共有した人だけが開ける
  FILE_LINK_SHARING: true,

  SHEET_NAME: '決裁申請',
  FOLDER_NAME: '決裁申請_添付ファイル'
};
/* ===================================================== */

const STATUS = {
  HONBU:   '本部確認待ち',
  RIJICHO: '理事長確認待ち',
  DONE:    '決裁済',
  BACK:    '差戻し',
  REJECT:  '却下'
};

// スプレッドシートの列（見出し名で読み書きするので、列の並べ替えは可。見出し名は変えないこと）
const HEADERS = [
  '申請No', 'ステータス', '受付日時', '最終更新',
  '起案日', '起案所属名', '起案者名', '園長メール',
  '件名', '詳細理由', '想定される効果', '時期',
  '業者1', '業者1金額', '業者2', '業者2金額', '業者3', '業者3金額', '採用業者',
  '金額', '予算', '添付ファイル',
  '本部判断', '本部承認者', '本部日時', '本部コメント',
  '理事長判断', '理事長名', '理事長日時', '理事長コメント',
  '再申請回数', '履歴', '編集キー'
];

/* ======================== 入口 ======================== */

// GET：件数（本部ダッシュボード用）。中身は返さない
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'count';
  try {
    if (action === 'count') return json_(countPending_());
    return json_({ ok: false, error: '不明な操作です' });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

// POST：申請・再申請・一覧・承認操作
function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'データを読み取れませんでした' });
  }
  const action = body.action || 'submit';   // 申請フォームは action なしで送ってくる
  try {
    switch (action) {
      case 'submit':   return json_(submit_(body));
      case 'load':     return json_(loadForResubmit_(body));
      case 'resubmit': return json_(resubmit_(body));
      case 'login':    return json_(login_(body));
      case 'list':     return json_(list_(body));
      case 'decide':   return json_(decide_(body));
      default:         return json_({ ok: false, error: '不明な操作です' });
    }
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

/* ======================== 申請（園長） ======================== */

function submit_(p) {
  validate_(p);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_();
    const appNo = nextAppNo_(sh);
    const now = new Date();
    const files = saveFiles_(appNo, p.files || [], 1);
    const rec = {};
    fillFromPayload_(rec, p);
    rec['申請No'] = appNo;
    rec['ステータス'] = STATUS.HONBU;
    rec['受付日時'] = now;
    rec['最終更新'] = now;
    rec['添付ファイル'] = JSON.stringify(files);
    rec['再申請回数'] = 0;
    rec['履歴'] = JSON.stringify([{ at: iso_(now), role: '園長', name: p.kianja, action: '申請', comment: '' }]);
    rec['編集キー'] = Utilities.getUuid();
    appendRecord_(sh, rec);
    SpreadsheetApp.flush();
    mailHonbu_(rec, false);
    return { ok: true, appNo: appNo };
  } finally {
    lock.releaseLock();
  }
}

// 差戻しされた申請を、修正用リンクから読み込む
function loadForResubmit_(p) {
  const found = findByNo_(p.appNo);
  if (!found) throw new Error('申請が見つかりません');
  const r = found.rec;
  if (r['編集キー'] !== p.editKey) throw new Error('このリンクは無効です');
  if (r['ステータス'] !== STATUS.BACK) throw new Error('この申請は現在「' + r['ステータス'] + '」のため修正できません');
  return { ok: true, item: publicItem_(r) };
}

function resubmit_(p) {
  validate_(p);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_();
    const found = findByNo_(p.appNo, sh);
    if (!found) throw new Error('申請が見つかりません');
    const r = found.rec;
    if (r['編集キー'] !== p.editKey) throw new Error('このリンクは無効です');
    if (r['ステータス'] !== STATUS.BACK) throw new Error('この申請は現在「' + r['ステータス'] + '」のため再申請できません');

    const now = new Date();
    const times = Number(r['再申請回数'] || 0) + 1;
    // 既存の添付は「残す」と指定されたものだけ残し、新しいファイルを追加
    const keep = Array.isArray(p.keepFiles) ? p.keepFiles : [];
    const oldFiles = parseJson_(r['添付ファイル'], []).filter(function (f) { return keep.indexOf(f.id) >= 0; });
    const newFiles = saveFiles_(p.appNo, p.files || [], times + 1);

    fillFromPayload_(r, p);
    r['ステータス'] = STATUS.HONBU;
    r['最終更新'] = now;
    r['添付ファイル'] = JSON.stringify(oldFiles.concat(newFiles));
    r['本部判断'] = ''; r['本部承認者'] = ''; r['本部日時'] = ''; r['本部コメント'] = '';
    r['理事長判断'] = ''; r['理事長名'] = ''; r['理事長日時'] = ''; r['理事長コメント'] = '';
    r['再申請回数'] = times;
    const hist = parseJson_(r['履歴'], []);
    hist.push({ at: iso_(now), role: '園長', name: p.kianja, action: '再申請', comment: '' });
    r['履歴'] = JSON.stringify(hist);
    writeRecord_(sh, found.row, r);
    SpreadsheetApp.flush();
    mailHonbu_(r, true);
    return { ok: true, appNo: p.appNo };
  } finally {
    lock.releaseLock();
  }
}

function validate_(p) {
  const need = { kian_date: '起案日', shozoku: '起案所属名', kianja: '起案者名', email: '連絡用メールアドレス', kenmei: '件名', riyu: '詳細理由', kingaku: '金額', budget: '予算' };
  Object.keys(need).forEach(function (k) {
    if (!String(p[k] || '').trim()) throw new Error(need[k] + 'が未入力です');
  });
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(p.email).trim())) throw new Error('メールアドレスの形式が正しくありません');
}

function fillFromPayload_(rec, p) {
  const v = p.vendors || [];
  rec['起案日'] = p.kian_date || '';
  rec['起案所属名'] = p.shozoku || '';
  rec['起案者名'] = p.kianja || '';
  rec['園長メール'] = String(p.email || '').trim();
  rec['件名'] = p.kenmei || '';
  rec['詳細理由'] = p.riyu || '';
  rec['想定される効果'] = p.kouka || '';
  rec['時期'] = p.jiki || '';
  for (let i = 0; i < 3; i++) {
    rec['業者' + (i + 1)] = (v[i] && v[i].name) || '';
    rec['業者' + (i + 1) + '金額'] = (v[i] && v[i].amt) || '';
  }
  rec['採用業者'] = p.adopt || '';
  rec['金額'] = p.kingaku || '';
  rec['予算'] = p.budget || '';
}

/* ======================== 承認（本部・理事長） ======================== */

const NEED_LOGIN = '本部ダッシュボードのログインが切れています。本部ダッシュボードからログインし直してください。';

// 本部ダッシュボードの通行証を確認し、押せる承認欄（honbu／rijicho）を返す。だめなら ''
//   確認結果は5分間だけ覚えておく（毎回問い合わせないように）
function roleOf_(token) {
  token = String(token || '');
  if (token.length < 40) return '';
  const cacheKey = 'auth_' + Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, token, Utilities.Charset.UTF_8));
  const cache = CacheService.getScriptCache();
  const hit = cache.get(cacheKey);
  if (hit !== null) return hit;

  let role = '';
  try {
    const res = UrlFetchApp.fetch(CONFIG.AUTH_URL + '?action=verify&token=' + encodeURIComponent(token),
      { muteHttpExceptions: true, followRedirects: true });
    const data = JSON.parse(res.getContentText());
    if (data && data.ok && data.id) role = CONFIG.ROLE_BY_ID[data.id] || '';
  } catch (err) {
    console.error('ログイン確認に失敗: ' + err);
    return '';   // 失敗は覚えない（次回もう一度問い合わせる）
  }
  cache.put(cacheKey, role, 300);
  return role;
}

function login_(p) {
  const role = roleOf_(p.token);
  if (!role) return { ok: false, error: NEED_LOGIN, needLogin: true };
  return { ok: true, role: role };
}

function list_(p) {
  const role = roleOf_(p.token);
  if (!role) return { ok: false, error: NEED_LOGIN, needLogin: true };
  const sh = sheet_();
  const items = readAll_(sh).map(function (x) { return publicItem_(x.rec); });
  items.sort(function (a, b) { return String(b.receivedAt).localeCompare(String(a.receivedAt)); });
  return { ok: true, role: role, items: items };
}

function decide_(p) {
  const role = roleOf_(p.token);
  if (!role) return { ok: false, error: NEED_LOGIN, needLogin: true };
  const decision = p.decision;
  if (['承認', '差戻し', '却下'].indexOf(decision) < 0) throw new Error('操作の種類が不正です');
  const name = String(p.name || '').trim();
  if (!name) throw new Error('承認者の氏名を入力してください');
  const comment = String(p.comment || '').trim();
  if (decision !== '承認' && !comment) throw new Error(decision + 'の理由を入力してください');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_();
    const found = findByNo_(p.appNo, sh);
    if (!found) throw new Error('申請が見つかりません');
    const r = found.rec;
    const expect = role === 'honbu' ? STATUS.HONBU : STATUS.RIJICHO;
    if (r['ステータス'] !== expect) {
      throw new Error('この申請は現在「' + r['ステータス'] + '」のため、' + (role === 'honbu' ? '法人本部' : '理事長') + 'の操作はできません（画面を更新してください）');
    }
    const now = new Date();
    const label = role === 'honbu' ? '法人本部' : '理事長';
    if (role === 'honbu') {
      r['本部判断'] = decision; r['本部承認者'] = name; r['本部日時'] = now; r['本部コメント'] = comment;
    } else {
      r['理事長判断'] = decision; r['理事長名'] = name; r['理事長日時'] = now; r['理事長コメント'] = comment;
    }
    if (decision === '承認') r['ステータス'] = role === 'honbu' ? STATUS.RIJICHO : STATUS.DONE;
    else if (decision === '差戻し') r['ステータス'] = STATUS.BACK;
    else r['ステータス'] = STATUS.REJECT;
    r['最終更新'] = now;
    const hist = parseJson_(r['履歴'], []);
    hist.push({ at: iso_(now), role: label, name: name, action: decision, comment: comment });
    r['履歴'] = JSON.stringify(hist);
    writeRecord_(sh, found.row, r);
    SpreadsheetApp.flush();

    // 結果を園長へ（本部が承認して理事長へ回すときは通知しない）
    if (r['ステータス'] === STATUS.DONE || r['ステータス'] === STATUS.BACK || r['ステータス'] === STATUS.REJECT) {
      mailResult_(r, label, decision, comment);
    }
    return { ok: true, item: publicItem_(r) };
  } finally {
    lock.releaseLock();
  }
}

/* ======================== 件数 ======================== */

function countPending_() {
  const all = readAll_(sheet_());
  let honbu = 0, rijicho = 0, back = 0;
  all.forEach(function (x) {
    const s = x.rec['ステータス'];
    if (s === STATUS.HONBU) honbu++;
    else if (s === STATUS.RIJICHO) rijicho++;
    else if (s === STATUS.BACK) back++;
  });
  // kessai：本部ダッシュボードの「決裁承認」カードに出す件数（本部＋理事長の未処理）
  return { ok: true, counts: { kessai: honbu + rijicho }, detail: { honbu: honbu, rijicho: rijicho, back: back } };
}

/* ======================== メール ======================== */

function mailHonbu_(r, isResubmit) {
  if (!CONFIG.HONBU_EMAIL) return;
  const subject = '【決裁申請' + (isResubmit ? '・再申請' : '') + '】' + r['起案所属名'] + '「' + r['件名'] + '」' + r['申請No'];
  const body = [
    '法人本部 ご担当者様',
    '',
    '決裁申請書が' + (isResubmit ? '修正のうえ再申請' : '申請') + 'されました。承認画面からご確認ください。',
    '',
    '申請No　：' + r['申請No'],
    '所属　　：' + r['起案所属名'],
    '起案者　：' + r['起案者名'],
    '件名　　：' + r['件名'],
    '金額　　：¥' + yen_(r['金額']) + '（' + r['予算'] + '）',
    '',
    '承認画面：' + CONFIG.APPROVE_URL + '?no=' + encodeURIComponent(r['申請No']),
    '',
    '※このメールは決裁承認システムから自動送信しています。'
  ].join('\n');
  safeMail_(CONFIG.HONBU_EMAIL, subject, body);
}

function mailResult_(r, label, decision, comment) {
  const to = r['園長メール'];
  if (!to) return;
  const status = r['ステータス'];
  const subject = '【決裁申請・' + status + '】「' + r['件名'] + '」' + r['申請No'];
  const lines = [
    r['起案者名'] + ' 様',
    '',
    'ご申請いただいた決裁申請書について、' + label + 'より「' + decision + '」となりました。',
    '',
    '申請No：' + r['申請No'],
    '件名　：' + r['件名'],
    '金額　：¥' + yen_(r['金額']),
    '結果　：' + status
  ];
  if (comment) lines.push('', '■ ' + label + 'からのコメント', comment);
  if (status === STATUS.DONE) {
    lines.push('', '法人本部・理事長の承認がそろい、決裁されました。');
  } else if (status === STATUS.BACK) {
    lines.push('', '内容を修正のうえ、下のリンクから再申請してください（このリンクはあなた専用です。他の人に転送しないでください）。',
      CONFIG.FORM_URL + '?resubmit=' + encodeURIComponent(r['申請No']) + '&ek=' + encodeURIComponent(r['編集キー']));
  } else if (status === STATUS.REJECT) {
    lines.push('', 'この申請は却下されたため、手続きは終了しました。');
  }
  lines.push('', '※このメールは決裁承認システムから自動送信しています。');
  safeMail_(to, subject, lines.join('\n'));
}

function safeMail_(to, subject, body) {
  try {
    MailApp.sendEmail({ to: to, subject: subject, body: body, name: 'しおどめグループ 決裁承認' });
  } catch (err) {
    console.error('メール送信失敗: ' + err);   // メールが失敗しても申請・承認自体は保存済み
  }
}

/* ======================== 添付ファイル ======================== */

function saveFiles_(appNo, files, round) {
  if (!files.length) return [];
  const root = rootFolder_();
  const it = root.getFoldersByName(appNo);
  const folder = it.hasNext() ? it.next() : root.createFolder(appNo);
  return files.map(function (f) {
    const blob = Utilities.newBlob(Utilities.base64Decode(f.dataB64), f.type || 'application/octet-stream', f.name);
    const file = folder.createFile(blob);
    if (CONFIG.FILE_LINK_SHARING) {
      try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) { console.warn(e); }
    }
    return { id: file.getId(), name: f.name, url: file.getUrl(), round: round };
  });
}

function rootFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('FOLDER_ID');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) { /* 削除されていたら作り直す */ }
  }
  const folder = DriveApp.createFolder(CONFIG.FOLDER_NAME);
  props.setProperty('FOLDER_ID', folder.getId());
  return folder;
}

/* ======================== シート操作 ======================== */

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.SHEET_NAME);
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold').setBackground('#e8eefc');
    sh.setFrozenRows(1);
  }
  return sh;
}

function headerMap_(sh) {
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const map = {};
  head.forEach(function (h, i) { if (h) map[h] = i; });
  // 足りない見出しがあれば右端に追加
  HEADERS.forEach(function (h) {
    if (!(h in map)) {
      const col = sh.getLastColumn() + 1;
      sh.getRange(1, col).setValue(h).setFontWeight('bold').setBackground('#e8eefc');
      map[h] = col - 1;
    }
  });
  return map;
}

function readAll_(sh) {
  const map = headerMap_(sh);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const width = sh.getLastColumn();
  const values = sh.getRange(2, 1, last - 1, width).getValues();
  const out = [];
  values.forEach(function (row, i) {
    if (!row[map['申請No']]) return;
    const rec = {};
    Object.keys(map).forEach(function (h) { rec[h] = row[map[h]]; });
    out.push({ row: i + 2, rec: rec });
  });
  return out;
}

function findByNo_(appNo, sh) {
  if (!appNo) return null;
  const all = readAll_(sh || sheet_());
  for (let i = 0; i < all.length; i++) if (String(all[i].rec['申請No']) === String(appNo)) return all[i];
  return null;
}

function appendRecord_(sh, rec) {
  const map = headerMap_(sh);
  const row = new Array(sh.getLastColumn()).fill('');
  Object.keys(rec).forEach(function (h) { if (h in map) row[map[h]] = rec[h]; });
  sh.appendRow(row);
}

function writeRecord_(sh, rowNo, rec) {
  const map = headerMap_(sh);
  const width = sh.getLastColumn();
  const row = sh.getRange(rowNo, 1, 1, width).getValues()[0];
  Object.keys(rec).forEach(function (h) { if (h in map) row[map[h]] = rec[h]; });
  sh.getRange(rowNo, 1, 1, width).setValues([row]);
}

// 申請No：K-YYMM-001 形式（月ごとに連番）
function nextAppNo_(sh) {
  const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
  const prefix = 'K-' + Utilities.formatDate(new Date(), tz, 'yyMM') + '-';
  let max = 0;
  readAll_(sh).forEach(function (x) {
    const no = String(x.rec['申請No']);
    if (no.indexOf(prefix) === 0) max = Math.max(max, Number(no.slice(prefix.length)) || 0);
  });
  return prefix + ('00' + (max + 1)).slice(-3);
}

// 画面に返す形（編集キーは返さない）
function publicItem_(r) {
  return {
    appNo: r['申請No'],
    status: r['ステータス'],
    receivedAt: iso_(r['受付日時']),
    updatedAt: iso_(r['最終更新']),
    kian_date: dateStr_(r['起案日']),
    shozoku: r['起案所属名'],
    kianja: r['起案者名'],
    email: r['園長メール'],
    kenmei: r['件名'],
    riyu: r['詳細理由'],
    kouka: r['想定される効果'],
    jiki: r['時期'],
    vendors: [1, 2, 3].map(function (i) { return { name: String(r['業者' + i] || ''), amt: String(r['業者' + i + '金額'] || '') }; }),
    adopt: String(r['採用業者'] || ''),
    kingaku: String(r['金額'] || ''),
    budget: r['予算'],
    files: parseJson_(r['添付ファイル'], []),
    honbu: { decision: r['本部判断'], name: r['本部承認者'], at: iso_(r['本部日時']), comment: r['本部コメント'] },
    rijicho: { decision: r['理事長判断'], name: r['理事長名'], at: iso_(r['理事長日時']), comment: r['理事長コメント'] },
    resubmitCount: Number(r['再申請回数'] || 0),
    history: parseJson_(r['履歴'], [])
  };
}

/* ======================== 小道具 ======================== */

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function yen_(v) {
  const d = String(v == null ? '' : v).replace(/[^\d]/g, '');
  return d ? Number(d).toLocaleString('ja-JP') : '0';
}
function parseJson_(s, fallback) {
  try { return s ? JSON.parse(s) : fallback; } catch (e) { return fallback; }
}
function iso_(d) {
  if (!d) return '';
  if (d instanceof Date) return Utilities.formatDate(d, 'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm:ss");
  return String(d);
}
function dateStr_(d) {
  if (!d) return '';
  if (d instanceof Date) return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
  return String(d);
}

/* ======================== 初回セットアップ ======================== */

// 最初に1回だけ実行：シート作成・添付フォルダ作成・メール/ドライブの権限許可
function setup() {
  sheet_();
  const folder = rootFolder_();
  console.log('添付フォルダ：' + folder.getUrl());
  console.log('今日あと送れるメール数：' + MailApp.getRemainingDailyQuota());
  UrlFetchApp.fetch(CONFIG.AUTH_URL + '?action=verify', { muteHttpExceptions: true });   // 外部への問い合わせの権限を許可するため
  console.log('準備OK。次は「デプロイ → 新しいデプロイ → ウェブアプリ」です。');
}