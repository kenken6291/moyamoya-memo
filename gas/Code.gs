/**
 * もやもやメモ ― 「思い出せない」記録＆AI分析 バックエンド（Google Apps Script）
 *
 * スクリプトプロパティ（プロジェクトの設定 > スクリプト プロパティ）
 *   GEMINI_API_KEY   : 必須。Google AI Studio で発行した API キー
 *   GEMINI_MODEL     : 任意。既定 gemini-2.5-flash（マルチモーダル対応モデルを指定）
 *   SPREADSHEET_ID   : 任意。未設定ならこのスクリプトが紐づくスプレッドシートを使用
 *   ROOT_FOLDER_ID   : 任意。未設定なら setup() 実行時に「もやもやメモ」フォルダを自動作成
 *
 * 初回は setup() をエディタから1回実行して権限承認とシート・フォルダ作成を行ってください。
 */

const APP_NAME = 'もやもやメモ';
// 更新のたびに書き換えると、GAS URL を開いたときに反映を確認できます
const APP_VERSION = '2026-10-06-community';
const TZ = 'Asia/Tokyo';
const SESSION_DAYS = 30;
const MAX_IMAGES = 4;
const MAX_IMAGE_BASE64 = 3 * 1024 * 1024; // 1枚あたり約3MB（base64）まで
const LOCK_FAILS = 5;
const LOCK_MINUTES = 15;

const CATEGORIES = ['人名', 'エンタメ', '日用品・モノ', '場所・地名', '言葉・用語', '予定・タスク', 'その他'];
const SITUATIONS = ['会話中', 'テレビ・動画', '本・雑誌', '仕事中', '外出・旅先', '家事中', 'その他'];

// 公開設定・みんなの反応
const VIS_PUBLIC = '公開';
const VIS_PRIVATE = '非公開';
const REACTIONS = { aruaru: 'それ、あるある', like: 'いいね' };
const COMMENT_MAX = 500;
const PUBLIC_PAGE = 20;

// [プログラム内キー, シート見出し]
const USER_COLS = [
  ['userId', 'ユーザーID'],
  ['email', 'メールアドレス'],
  ['displayName', '表示名'],
  ['passwordHash', 'パスワードハッシュ'],
  ['salt', 'ソルト'],
  ['mustChangePassword', '要パスワード変更'],
  ['failedCount', 'ログイン失敗回数'],
  ['lockedUntil', 'ロック解除日時'],
  ['createdAt', '登録日時'],
  ['lastLoginAt', '最終ログイン日時'],
];

const RECORD_COLS = [
  ['recordId', '記録ID'],
  ['userId', 'ユーザーID'],
  ['createdAt', '記録日時'],
  ['occurredAt', '発生日時'],
  ['situation', '状況'],
  ['clues', '手がかり'],
  ['voiceRaw', '音声原文'],
  ['inputMethods', '入力手段'],
  ['imageUrls', '画像URL'],
  ['imageFileIds', '画像ファイルID'],
  ['imageKinds', '画像種別'],
  ['extractedKeywords', '抽出キーワード'],
  ['aiInterpretation', 'AI解釈'],
  ['candidates', 'AI候補(JSON)'],
  ['answer', '正解'],
  ['category', 'カテゴリ'],
  ['subCategory', 'サブカテゴリ'],
  ['tags', 'タグ'],
  ['memoryHint', '覚え方ヒント'],
  ['resolvedBy', '解決方法'],
  ['resolvedAt', '解決日時'],
  ['resolveMinutes', '解決までの分数'],
  ['status', 'ステータス'],
  ['updatedAt', '更新日時'],
  ['visibility', '公開設定'],
  ['publishedAt', '公開日時'],
];

const COMMENT_COLS = [
  ['commentId', 'コメントID'],
  ['recordId', '記録ID'],
  ['userId', 'ユーザーID'],
  ['text', '本文'],
  ['createdAt', '投稿日時'],
  ['updatedAt', '更新日時'],
  ['editedBy', '最終編集者ID'],
];

const REACTION_COLS = [
  ['recordId', '記録ID'],
  ['userId', 'ユーザーID'],
  ['type', '種類'],
  ['createdAt', '日時'],
];

const SESSION_COLS = [
  ['token', 'トークン'],
  ['userId', 'ユーザーID'],
  ['expiresAt', '有効期限'],
];

/* =========================================================
 * エントリポイント
 * ======================================================= */

function doGet() {
  return json_({ ok: true, app: APP_NAME, version: APP_VERSION, time: new Date().toISOString() });
}

function doPost(e) {
  let body = {};
  try {
    body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return json_({ ok: false, error: 'リクエストの形式が正しくありません。' });
  }
  const action = body.action;
  const routes = {
    register: register_,
    login: login_,
    logout: logout_,
    me: me_,
    forgotPassword: forgotPassword_,
    changePassword: changePassword_,
    cleanVoice: cleanVoice_,
    infer: infer_,
    refine: refine_,
    resolve: resolve_,
    updateRecord: updateRecord_,
    reclassify: reclassify_,
    listRecords: listRecords_,
    deleteRecord: deleteRecord_,
    getImage: getImage_,
    exportToDrive: exportToDrive_,
    setVisibility: setVisibility_,
    listPublic: listPublic_,
    getComments: getComments_,
    addComment: addComment_,
    editComment: editComment_,
    deleteComment: deleteComment_,
    toggleReaction: toggleReaction_,
  };
  try {
    if (!routes[action]) throw new AppError('不明な操作です: ' + action);
    const result = routes[action](body) || {};
    return json_(Object.assign({ ok: true }, result));
  } catch (err) {
    if (err instanceof AppError) return json_({ ok: false, error: err.message, code: err.code || '' });
    console.error(err && err.stack ? err.stack : err);
    return json_({ ok: false, error: 'サーバーでエラーが発生しました: ' + (err && err.message ? err.message : err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

class AppError extends Error {
  constructor(message, code) {
    super(message);
    this.code = code;
  }
}

/* =========================================================
 * 初期セットアップ（エディタから1回実行）
 * ======================================================= */

function setup() {
  sheet_('Users', USER_COLS);
  sheet_('Records', RECORD_COLS);
  sheet_('Sessions', SESSION_COLS);
  sheet_('Comments', COMMENT_COLS);
  sheet_('Reactions', REACTION_COLS);
  const root = rootFolder_();
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('GEMINI_API_KEY')) {
    console.warn('GEMINI_API_KEY が未設定です。スクリプトプロパティに設定してください。');
  }
  console.log('スプレッドシート: ' + ss_().getUrl());
  console.log('Driveフォルダ: ' + root.getUrl());
  // メール送信の権限承認用（自分宛には送らない）
  MailApp.getRemainingDailyQuota();
}

/* =========================================================
 * スプレッドシート共通処理（見出し名ベースで列を解決）
 * ======================================================= */

function ss_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  const ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new AppError('スプレッドシートが見つかりません。SPREADSHEET_ID を設定してください。');
  return ss;
}

/** シートを取得。無ければ作成し、足りない見出しは右端に追加する（既存データの列追加移行に対応） */
function sheet_(name, cols) {
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, cols.length).setValues([cols.map(c => c[1])]).setFontWeight('bold');
    sh.setFrozenRows(1);
    return sh;
  }
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const missing = cols.map(c => c[1]).filter(h => header.indexOf(h) === -1);
  if (missing.length) {
    const start = header.filter(h => h !== '').length + 1;
    sh.getRange(1, start, 1, missing.length).setValues([missing]).setFontWeight('bold');
  }
  return sh;
}

function colIndex_(sh, cols) {
  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const idx = {};
  cols.forEach(([key, label]) => { idx[key] = header.indexOf(label); });
  return { idx, width: header.length };
}

function readAll_(name, cols) {
  const sh = sheet_(name, cols);
  const { idx } = colIndex_(sh, cols);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const values = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  return values.map((row, i) => {
    const o = { _row: i + 2 };
    cols.forEach(([key]) => { o[key] = idx[key] >= 0 ? row[idx[key]] : ''; });
    return o;
  });
}

function append_(name, cols, obj) {
  const sh = sheet_(name, cols);
  const { idx, width } = colIndex_(sh, cols);
  const row = new Array(width).fill('');
  cols.forEach(([key]) => { if (idx[key] >= 0 && obj[key] !== undefined) row[idx[key]] = obj[key]; });
  sh.appendRow(row);
}

function update_(name, cols, rowNum, patch) {
  const sh = sheet_(name, cols);
  const { idx } = colIndex_(sh, cols);
  Object.keys(patch).forEach(key => {
    if (idx[key] >= 0) sh.getRange(rowNum, idx[key] + 1).setValue(patch[key]);
  });
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function toIso_(v) {
  if (v instanceof Date) return v.toISOString();
  return v ? String(v) : '';
}

function fmt_(d) {
  return d instanceof Date ? Utilities.formatDate(d, TZ, 'yyyy/MM/dd HH:mm') : String(d || '');
}

/* =========================================================
 * 認証
 * ======================================================= */

function hash_(password, salt) {
  let h = salt + '::' + password;
  for (let i = 0; i < 200; i++) {
    const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h, Utilities.Charset.UTF_8);
    h = Utilities.base64Encode(bytes);
  }
  return h;
}

function randomPassword_() {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789';
  let s = '';
  for (let i = 0; i < 10; i++) s += chars.charAt(Math.floor(Math.random() * chars.length));
  return s;
}

function normEmail_(email) {
  return String(email || '').trim().toLowerCase();
}

function findUserByEmail_(email) {
  const e = normEmail_(email);
  return readAll_('Users', USER_COLS).find(u => normEmail_(u.email) === e) || null;
}

function findUserById_(userId) {
  return readAll_('Users', USER_COLS).find(u => u.userId === userId) || null;
}

function publicUser_(u) {
  return {
    userId: u.userId,
    email: u.email,
    displayName: u.displayName,
    mustChangePassword: u.mustChangePassword === true || u.mustChangePassword === 'TRUE',
  };
}

function createSession_(userId) {
  const token = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
  const expiresAt = new Date(Date.now() + SESSION_DAYS * 86400000);
  append_('Sessions', SESSION_COLS, { token, userId, expiresAt });
  CacheService.getScriptCache().put('s_' + token, userId, 21600);
  return token;
}

/** トークンからユーザーを取得。無効なら AUTH エラー */
function requireUser_(body) {
  const token = String(body.token || '');
  if (!token) throw new AppError('ログインが必要です。', 'AUTH');
  const cache = CacheService.getScriptCache();
  let userId = cache.get('s_' + token);
  if (!userId) {
    const s = readAll_('Sessions', SESSION_COLS).find(r => r.token === token);
    if (!s || new Date(s.expiresAt).getTime() < Date.now()) {
      throw new AppError('ログインの有効期限が切れました。もう一度ログインしてください。', 'AUTH');
    }
    userId = s.userId;
    cache.put('s_' + token, userId, 21600);
  }
  const user = findUserById_(userId);
  if (!user) throw new AppError('ユーザーが見つかりません。', 'AUTH');
  return user;
}

/**
 * 新規登録：メールアドレス＋表示名のみ受け付け、仮パスワードをメールで送る。
 * 仮パスワードでログインすると、パスワード変更が必須になる（本人のメールであることの確認を兼ねる）。
 */
function register_(body) {
  const email = normEmail_(body.email);
  const displayName = String(body.displayName || '').trim();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new AppError('メールアドレスの形式が正しくありません。');
  if (!displayName) throw new AppError('表示名を入力してください。');
  if (displayName.length > 30) throw new AppError('表示名は30文字以内にしてください。');

  const temp = randomPassword_();
  const user = withLock_(() => {
    if (findUserByEmail_(email)) {
      throw new AppError('このメールアドレスは登録済みです。パスワードがわからない場合は「パスワードをお忘れの方」から再発行してください。');
    }
    const salt = Utilities.getUuid();
    const u = {
      userId: 'U' + Utilities.getUuid().replace(/-/g, '').slice(0, 12),
      email,
      displayName,
      passwordHash: hash_(temp, salt),
      salt,
      mustChangePassword: true,
      failedCount: 0,
      lockedUntil: '',
      createdAt: new Date(),
      lastLoginAt: '',
    };
    append_('Users', USER_COLS, u);
    return u;
  });

  try {
    MailApp.sendEmail({
      to: email,
      subject: `【${APP_NAME}】仮パスワードのお知らせ（ご登録ありがとうございます）`,
      body: `${displayName} さん\n\n${APP_NAME} へのご登録ありがとうございます。\n\n` +
        `ログイン用メールアドレス: ${email}\n` +
        `仮パスワード: ${temp}\n\n` +
        'アプリの「ログイン」から上記でログインすると、新しいパスワードの設定画面が表示されます。\n' +
        'ご自身で決めたパスワードに変更してからご利用ください。\n\n' +
        '「あれ、なんだっけ？」と思ったら、文字・声・手書き・写真のどれでも気軽に記録してください。\n\n' +
        '※このメールに心当たりがない場合は破棄してください。',
    });
  } catch (err) {
    // メールが届かないとログインできないため、登録を取り消す
    withLock_(() => {
      const u = findUserById_(user.userId);
      if (u) sheet_('Users', USER_COLS).deleteRow(u._row);
    });
    throw new AppError('仮パスワードのメールを送れませんでした。メールアドレスを確認して、もう一度登録してください。');
  }

  return { email, message: `${email} に仮パスワードを送りました。メールを確認してログインしてください。` };
}

function login_(body) {
  const email = normEmail_(body.email);
  const password = String(body.password || '');
  const u = findUserByEmail_(email);
  if (!u) throw new AppError('メールアドレスまたはパスワードが違います。');
  if (u.lockedUntil && new Date(u.lockedUntil).getTime() > Date.now()) {
    throw new AppError(`ログインに${LOCK_FAILS}回失敗したため一時的にロック中です。${fmt_(new Date(u.lockedUntil))} 以降に再度お試しください。`);
  }
  if (hash_(password, u.salt) !== u.passwordHash) {
    const fails = Number(u.failedCount || 0) + 1;
    const patch = { failedCount: fails };
    if (fails >= LOCK_FAILS) {
      patch.lockedUntil = new Date(Date.now() + LOCK_MINUTES * 60000);
      patch.failedCount = 0;
    }
    withLock_(() => update_('Users', USER_COLS, u._row, patch));
    throw new AppError('メールアドレスまたはパスワードが違います。');
  }
  withLock_(() => update_('Users', USER_COLS, u._row, { failedCount: 0, lockedUntil: '', lastLoginAt: new Date() }));
  return { token: createSession_(u.userId), user: publicUser_(u) };
}

function logout_(body) {
  const token = String(body.token || '');
  if (!token) return {};
  CacheService.getScriptCache().remove('s_' + token);
  withLock_(() => {
    const s = readAll_('Sessions', SESSION_COLS).find(r => r.token === token);
    if (s) sheet_('Sessions', SESSION_COLS).deleteRow(s._row);
  });
  return {};
}

function me_(body) {
  return { user: publicUser_(requireUser_(body)) };
}

function forgotPassword_(body) {
  const email = normEmail_(body.email);
  const message = '登録済みのメールアドレスであれば、仮パスワードを送信しました。メールをご確認ください。';
  const u = findUserByEmail_(email);
  if (!u) return { message };
  const temp = randomPassword_();
  const salt = Utilities.getUuid();
  withLock_(() => update_('Users', USER_COLS, u._row, {
    passwordHash: hash_(temp, salt),
    salt,
    mustChangePassword: true,
    failedCount: 0,
    lockedUntil: '',
  }));
  MailApp.sendEmail({
    to: u.email,
    subject: `【${APP_NAME}】仮パスワードのお知らせ`,
    body: `${u.displayName} さん\n\nパスワード再発行のお申し込みを受け付けました。\n\n` +
      `仮パスワード: ${temp}\n\n` +
      'この仮パスワードでログインすると、新しいパスワードの設定画面が表示されます。\n' +
      '※お申し込みに心当たりがない場合は、このメールを破棄してください（ただし仮パスワードに切り替わっています）。',
  });
  return { message };
}

function changePassword_(body) {
  const u = requireUser_(body);
  const current = String(body.currentPassword || '');
  const next = String(body.newPassword || '');
  if (hash_(current, u.salt) !== u.passwordHash) throw new AppError('現在のパスワードが違います。');
  if (next.length < 8) throw new AppError('新しいパスワードは8文字以上にしてください。');
  if (next === current) throw new AppError('新しいパスワードは現在と別のものにしてください。');
  const salt = Utilities.getUuid();
  withLock_(() => update_('Users', USER_COLS, u._row, {
    passwordHash: hash_(next, salt),
    salt,
    mustChangePassword: false,
  }));
  return { message: 'パスワードを変更しました。' };
}

/* =========================================================
 * Google Drive
 * ======================================================= */

function rootFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('ROOT_FOLDER_ID');
  if (id) return DriveApp.getFolderById(id);
  const folder = DriveApp.createFolder(APP_NAME);
  props.setProperty('ROOT_FOLDER_ID', folder.getId());
  return folder;
}

function childFolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

/** ユーザー専用フォルダ（ルート/users/表示名_ID/サブ） */
function userFolder_(user, sub) {
  const users = childFolder_(rootFolder_(), 'users');
  const safeName = String(user.displayName).replace(/[\\/:*?"<>|]/g, '_');
  const mine = childFolder_(users, `${safeName}_${user.userId}`);
  return sub ? childFolder_(mine, sub) : mine;
}

/** base64画像をDriveへ保存 */
function saveImage_(user, img, recordId, n) {
  const mimeType = img.mimeType === 'image/png' ? 'image/png' : 'image/jpeg';
  const ext = mimeType === 'image/png' ? 'png' : 'jpg';
  const kindLabel = img.kind === 'handwriting' ? '手書き' : '写真';
  const stamp = Utilities.formatDate(new Date(), TZ, 'yyyyMMdd_HHmmss');
  const blob = Utilities.newBlob(Utilities.base64Decode(img.data), mimeType, `${stamp}_${recordId}_${kindLabel}${n}.${ext}`);
  const file = userFolder_(user, 'images').createFile(blob);
  file.setDescription(`${APP_NAME} 記録ID: ${recordId}`);
  return { id: file.getId(), url: file.getUrl(), kind: img.kind };
}

/* =========================================================
 * Gemini
 * ======================================================= */

function callGemini_(parts, systemText, temperature) {
  const props = PropertiesService.getScriptProperties();
  // 貼り付け時に混入しがちな空白・改行・引用符を除去
  const key = String(props.getProperty('GEMINI_API_KEY') || '').replace(/[\s"'“”‘’]/g, '');
  if (!key) throw new AppError('GEMINI_API_KEY が設定されていません（GASのスクリプトプロパティ）。');
  const model = props.getProperty('GEMINI_MODEL') || 'gemini-2.5-flash';
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent`;
  const payload = {
    systemInstruction: { parts: [{ text: systemText }] },
    contents: [{ role: 'user', parts }],
    generationConfig: {
      temperature: temperature == null ? 0.4 : temperature,
      responseMimeType: 'application/json',
    },
  };
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-goog-api-key': key },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  let res;
  for (let attempt = 0; attempt < 3; attempt++) {
    res = UrlFetchApp.fetch(url, options);
    const code = res.getResponseCode();
    if (code === 200) break;
    if (code === 429 || code >= 500) {
      Utilities.sleep(1500 * (attempt + 1));
      continue;
    }
    throw new AppError(geminiErrorMessage_(code, res.getContentText()));
  }
  if (res.getResponseCode() !== 200) throw new AppError('AIが混み合っています。少し時間をおいて再度お試しください。');

  const data = JSON.parse(res.getContentText());
  const cand = data.candidates && data.candidates[0];
  const text = cand && cand.content && cand.content.parts
    ? cand.content.parts.map(p => p.text || '').join('')
    : '';
  if (!text) throw new AppError('AIから応答がありませんでした（安全フィルタ等でブロックされた可能性があります）。');
  return parseJsonLoose_(text);
}

/** Gemini APIのエラーを利用者向けの日本語に変換 */
function geminiErrorMessage_(code, bodyText) {
  let reason = '', message = '';
  try {
    const err = JSON.parse(bodyText).error || {};
    message = err.message || '';
    (err.details || []).forEach(d => { if (d.reason) reason = d.reason; });
  } catch (e) { message = String(bodyText).slice(0, 200); }
  console.error(`Gemini API error ${code} ${reason}: ${message}`);
  const map = {
    API_KEY_INVALID: 'Gemini APIキーが無効です。GASのスクリプトプロパティ「GEMINI_API_KEY」を正しいキーに設定し直してください。',
    API_KEY_HTTP_REFERRER_BLOCKED: 'Gemini APIキーに「ウェブサイト制限」がかかっています。Google Cloudの認証情報でアプリケーションの制限を「なし」にしてください。',
    API_KEY_SERVICE_BLOCKED: 'このAPIキーではGemini APIが許可されていません。キーのAPI制限に「Generative Language API」を追加してください。',
    SERVICE_DISABLED: 'このキーのプロジェクトでGemini API（Generative Language API）が有効になっていません。Google AI Studioで発行したキーを使ってください。',
  };
  if (map[reason]) return map[reason];
  if (code === 404) return 'Geminiのモデル名が見つかりません。スクリプトプロパティ「GEMINI_MODEL」を確認してください（例：gemini-2.5-flash）。';
  if (code === 400 && /location is not supported/i.test(message)) return 'この地域からはGemini APIを利用できません。';
  if (code === 403) return 'Gemini APIへのアクセスが拒否されました（' + (reason || message.slice(0, 80)) + '）。';
  return `AIの呼び出しに失敗しました（${code}）。${message.slice(0, 120)}`;
}

/** エディタから実行して、APIキーとモデルの設定を確認する */
function testGemini() {
  const r = callGemini_([{ text: '「テスト成功」という文字列を ok キーに入れたJSONを返してください。' }], 'JSONのみを出力してください。', 0);
  console.log('Gemini接続OK: ' + JSON.stringify(r));
}

function parseJsonLoose_(text) {
  const cleaned = String(text).replace(/```json|```/g, '').trim();
  try {
    return JSON.parse(cleaned);
  } catch (e) {
    const m = cleaned.match(/\{[\s\S]*\}/);
    if (m) return JSON.parse(m[0]);
    throw new AppError('AIの応答を解析できませんでした。');
  }
}

const INFER_SYSTEM = [
  'あなたは「喉まで出かかっているのに思い出せない」人を助ける、推理が得意な記憶アシスタントです。',
  'ユーザーの断片的な手がかりから、思い出したい対象（人名・作品名・モノの名前・地名・言葉・予定など）を推論します。',
  '',
  '# 入力の性質（重要）',
  '- テキストは音声認識の書き起こし（AIで整理済み）を含むことがあります。整理前の原文が参考として付くこともあります。「えーっと」「ほら」「あの」「なんだっけ」「あれあれ」等のつなぎ言葉、言い直し、語順の乱れ、同音異義語の誤変換（例：「刑事」→「掲示」「経時」）が混ざります。つなぎ言葉は無視し、誤変換は文脈から正しい語に読み替えて、意味のある手がかりだけを抽出してください。',
  '- 「〜に出てた」「〜っぽい」「〜みたいな」など曖昧な表現も重要な手がかりです。年代感（「昔の」「昭和の」）、共演者、見た場所、形、色、音の響き（「〇で始まる」「3文字くらい」）を特に重視してください。',
  '- 画像が添付されることがあります。種別は各画像の直前に示します。',
  '  - 手書き: 文字なら正確に文字起こしし（崩れた字は最も妥当な読みを採用）、絵なら輪郭・形状・特徴（帽子の形、髪型、ロゴの形など）を言語化して推論に使ってください。',
  '  - 写真: 写っている物・植物・建物・看板・商品・画面の文字を読み取り、対象を特定してください。植物は和名、看板や商品は名称、風景は場所の候補を挙げてください。',
  '- 人物が写った写真について、顔立ちだけを根拠に実在の人物を特定しないでください。番組テロップ、雑誌の見出し、名札、キャプションなど画像内の文字や文脈、ユーザーのテキスト手がかりを根拠にしてください。それが無い場合は、髪型・服装・雰囲気などの特徴を記述し、候補は手がかりから推測できる範囲に留め、特定のために有効な追加の手がかりを質問してください。一般の個人（知人・通行人など）の身元は推測しないでください。',
  '',
  '# 出力',
  '次のJSONのみを出力してください（前置き・マークダウン不要）。',
  '{',
  '  "extractedKeywords": ["手がかりから抽出した意味のある語（5〜10個）"],',
  '  "imageAnalysis": [{"index": 1, "kind": "手書き|写真", "transcription": "読み取れた文字（無ければ空）", "description": "見えている内容・特徴の要約"}],',
  '  "interpretation": "ユーザーが思い出したい対象を一文で言い換えたもの",',
  '  "candidates": [',
  '    {"name": "候補の正式名称", "confidence": "高|中|低", "reason": "手がかりとの一致点（60字以内）", "category": "人名|エンタメ|日用品・モノ|場所・地名|言葉・用語|予定・タスク|その他"}',
  '  ],',
  '  "followUpQuestion": "候補を絞り込むために次に聞くと良い質問を1つ（自信が高ければ空文字）"',
  '}',
  'candidates は確からしい順に1〜3件。手がかりが乏しくても最低1件は推測を出し、confidence を「低」にしてください。',
  '予定・タスク（「何をしに来たんだっけ」「何か頼まれてた」等）の場合は、状況から考えられる行動を候補にしてください。',
  '画像が無い場合、imageAnalysis は空配列にしてください。',
].join('\n');

function buildInferParts_(rec, images, extraClues) {
  const parts = [];
  const lines = [
    `【発生日時】${fmt_(rec.occurredAt instanceof Date ? rec.occurredAt : new Date(rec.occurredAt))}`,
    `【状況】${rec.situation || '未指定'}`,
    `【入力手段】${rec.inputMethods || 'テキスト'}`,
    '【手がかり（口語・音声書き起こしを含む場合あり）】',
    rec.clues ? String(rec.clues) : '（テキストの手がかりなし。画像から推論してください）',
  ];
  if (rec.voiceRaw) {
    lines.push('【参考：音声認識の原文（整理前。上の手がかりで聞き落とした語があれば拾ってください）】');
    lines.push(String(rec.voiceRaw).slice(0, 2000));
  }
  if (extraClues) {
    lines.push('【追加の手がかり】');
    lines.push(extraClues);
  }
  if (rec.prevCandidates && rec.prevCandidates.length) {
    lines.push('【前回の候補（ユーザーはまだ正解と確定していません。より良い候補があれば入れ替えてください）】');
    lines.push(rec.prevCandidates.map(c => c.name).join('、'));
  }
  parts.push({ text: lines.join('\n') });
  images.forEach((img, i) => {
    parts.push({ text: `--- 画像${i + 1}（${img.kind === 'handwriting' ? '手書き' : '写真'}） ---` });
    parts.push({ inlineData: { mimeType: img.mimeType, data: img.data } });
  });
  return parts;
}

function normalizeInference_(ai) {
  const candidates = (Array.isArray(ai.candidates) ? ai.candidates : [])
    .filter(c => c && c.name)
    .slice(0, 3)
    .map(c => ({
      name: String(c.name),
      confidence: ['高', '中', '低'].indexOf(c.confidence) >= 0 ? c.confidence : '低',
      reason: String(c.reason || ''),
      category: CATEGORIES.indexOf(c.category) >= 0 ? c.category : 'その他',
    }));
  return {
    extractedKeywords: (Array.isArray(ai.extractedKeywords) ? ai.extractedKeywords : []).map(String).slice(0, 12),
    imageAnalysis: Array.isArray(ai.imageAnalysis) ? ai.imageAnalysis : [],
    interpretation: String(ai.interpretation || ''),
    candidates,
    followUpQuestion: String(ai.followUpQuestion || ''),
  };
}

const CLEAN_VOICE_SYSTEM = [
  'あなたは音声認識の書き起こしを整える編集者です。',
  'ユーザーは「思い出せない何か」の手がかりを口頭で話しています。書き起こしには次のノイズが含まれます。',
  '- 同じ語句の繰り返しや、途中経過の重複（例：「諫早 諫早出身の 諫早出身の体操の選手」→「諫早出身の体操の選手」）',
  '- つなぎ言葉（えーっと、あの、ほら、なんだっけ、あれあれ、うーん 等）や言い直し',
  '- 同音異義語・固有名詞の誤変換（例：刑事→掲示、地名や人名の漢字違い）',
  '',
  '# ルール',
  '- 重複とつなぎ言葉を取り除き、言い直しは最後の言い方を採用する。',
  '- 誤変換は文脈から明らかな場合だけ直す。自信がない固有名詞は聞こえたとおりに残す。',
  '- 話していない情報を足さない。答え（人名や作品名など）を推測して書き込まない。',
  '- 「〜だった気がする」「〜かも」など確信度を表す言い方は残す。',
  '- 手がかりとして読みやすい、短く自然な日本語の文にする（1〜3文程度）。箇条書きにしない。',
  '- 既存の入力内容が与えられた場合、それと重複する内容は省く。',
  '',
  '次のJSONのみを出力: {"cleaned": "整えた文"}',
].join('\n');

/** 音声書き起こしの整理（重複・フィラー除去、誤変換修正） */
function cleanVoice_(body) {
  requireUser_(body);
  const text = String(body.text || '').trim().slice(0, 3000);
  if (!text) throw new AppError('整理する音声テキストがありません。');
  const context = String(body.context || '').trim().slice(0, 1500);
  const prompt = (context ? `【すでに入力済みの内容】\n${context}\n\n` : '') + `【音声認識の書き起こし】\n${text}`;
  const ai = callGemini_([{ text: prompt }], CLEAN_VOICE_SYSTEM, 0.1);
  const cleaned = String(ai.cleaned || '').trim();
  return { cleaned: cleaned || text };
}

const CLASSIFY_SYSTEM = [
  'あなたは「思い出せなかったもの」を分類する整理係です。',
  '当初の手がかりと、確定した正解から、カテゴリ・サブカテゴリ・タグを決めてください。',
  `カテゴリは次から必ず1つ: ${CATEGORIES.join(' / ')}`,
  'サブカテゴリ例: 人名→俳優/歌手/著名人/歴史人物/知人、エンタメ→映画/ドラマ/音楽/本/アニメ/番組、日用品・モノ→道具/食べ物/家電/植物/ファッション、場所・地名→店/観光地/地名/建物、言葉・用語→慣用句/専門用語/外来語/漢字。',
  'また、次に同じものを思い出しやすくするための短い覚え方のヒント（語呂・連想・関連付け、40字以内）を1つ作ってください。',
  '次のJSONのみを出力: {"category": "", "subCategory": "", "tags": ["3〜5個"], "memoryHint": ""}',
].join('\n');

function classify_(clues, answer, situation) {
  const ai = callGemini_([{ text: `【手がかり】${clues || '（画像のみ）'}\n【状況】${situation || '未指定'}\n【正解】${answer}` }], CLASSIFY_SYSTEM, 0.2);
  return {
    category: CATEGORIES.indexOf(ai.category) >= 0 ? ai.category : 'その他',
    subCategory: String(ai.subCategory || ''),
    tags: (Array.isArray(ai.tags) ? ai.tags : []).map(String).slice(0, 6),
    memoryHint: String(ai.memoryHint || ''),
  };
}

/* =========================================================
 * 記録：推論・再推論・確定
 * ======================================================= */

function validateImages_(images) {
  if (!Array.isArray(images)) return [];
  if (images.length > MAX_IMAGES) throw new AppError(`画像は${MAX_IMAGES}枚までです。`);
  return images.map(img => {
    if (!img || !img.data) throw new AppError('画像データが空です。');
    if (String(img.data).length > MAX_IMAGE_BASE64) throw new AppError('画像サイズが大きすぎます。');
    return {
      kind: img.kind === 'handwriting' ? 'handwriting' : 'photo',
      mimeType: img.mimeType === 'image/png' ? 'image/png' : 'image/jpeg',
      data: String(img.data).replace(/^data:image\/\w+;base64,/, ''),
    };
  });
}

function infer_(body) {
  const user = requireUser_(body);
  const clues = String(body.clues || '').trim().slice(0, 2000);
  const images = validateImages_(body.images);
  if (!clues && images.length === 0) throw new AppError('手がかりを文字・声・手書き・写真のいずれかで入力してください。');

  const occurredAt = body.occurredAt ? new Date(body.occurredAt) : new Date();
  const methods = (Array.isArray(body.inputMethods) ? body.inputMethods : []).filter(m => ['テキスト', '音声', '手書き', '写真'].indexOf(m) >= 0);
  const recordId = 'R' + Utilities.formatDate(new Date(), TZ, 'yyMMddHHmmss') + Utilities.getUuid().slice(0, 4);

  // 1) 画像をDriveへ保存（AIが失敗しても写真は残す）
  const saved = images.map((img, i) => saveImage_(user, img, recordId, i + 1));

  const rec = {
    recordId,
    userId: user.userId,
    createdAt: new Date(),
    occurredAt: isNaN(occurredAt.getTime()) ? new Date() : occurredAt,
    situation: String(body.situation || '').slice(0, 50),
    clues,
    voiceRaw: String(body.voiceRaw || '').trim().slice(0, 3000),
    inputMethods: methods.join(',') || (clues ? 'テキスト' : ''),
    imageUrls: saved.map(s => s.url).join('\n'),
    imageFileIds: saved.map(s => s.id).join(','),
    imageKinds: saved.map(s => s.kind).join(','),
    status: '未解決',
    visibility: body.visibility === VIS_PUBLIC ? VIS_PUBLIC : VIS_PRIVATE,
    updatedAt: new Date(),
  };
  if (rec.visibility === VIS_PUBLIC) rec.publishedAt = new Date();

  // 2) Gemini推論
  let result = null;
  let aiError = '';
  try {
    result = normalizeInference_(callGemini_(buildInferParts_(rec, images), INFER_SYSTEM, 0.5));
  } catch (err) {
    aiError = err.message || String(err);
    result = { extractedKeywords: [], imageAnalysis: [], interpretation: '', candidates: [], followUpQuestion: '' };
  }
  rec.extractedKeywords = result.extractedKeywords.join('、');
  rec.aiInterpretation = result.interpretation;
  rec.candidates = JSON.stringify(result.candidates);

  // 3) 記録
  withLock_(() => append_('Records', RECORD_COLS, rec));

  return { record: serializeRecord_(rec), inference: result, aiError };
}

function findOwnRecord_(user, recordId) {
  const r = readAll_('Records', RECORD_COLS).find(x => x.recordId === recordId);
  if (!r || r.userId !== user.userId) throw new AppError('記録が見つかりません。');
  return r;
}

function imagesFromRecord_(r) {
  const ids = String(r.imageFileIds || '').split(',').filter(Boolean);
  const kinds = String(r.imageKinds || '').split(',');
  return ids.map((id, i) => {
    try {
      const blob = DriveApp.getFileById(id).getBlob();
      return { kind: kinds[i] || 'photo', mimeType: blob.getContentType(), data: Utilities.base64Encode(blob.getBytes()) };
    } catch (e) {
      return null;
    }
  }).filter(Boolean);
}

/** 手がかりを追加して同じ記録で再推論 */
function refine_(body) {
  const user = requireUser_(body);
  const r = findOwnRecord_(user, body.recordId);
  const extra = String(body.extraClues || '').trim().slice(0, 1000);
  if (!extra) throw new AppError('追加の手がかりを入力してください。');

  let prev = [];
  try { prev = JSON.parse(r.candidates || '[]'); } catch (e) { prev = []; }
  const images = imagesFromRecord_(r);
  const rec = Object.assign({}, r, { prevCandidates: prev });
  const addRaw = String(body.voiceRaw || '').trim();
  if (addRaw) rec.voiceRaw = (r.voiceRaw ? r.voiceRaw + '\n' : '') + addRaw;
  const result = normalizeInference_(callGemini_(buildInferParts_(rec, images, extra), INFER_SYSTEM, 0.5));

  const methods = String(r.inputMethods || '').split(',').filter(Boolean);
  (Array.isArray(body.inputMethods) ? body.inputMethods : []).forEach(m => { if (methods.indexOf(m) < 0) methods.push(m); });

  const extraRaw = String(body.voiceRaw || '').trim().slice(0, 2000);
  const patch = {
    clues: (r.clues ? r.clues + '\n' : '') + '＋ ' + extra,
    voiceRaw: extraRaw ? (r.voiceRaw ? r.voiceRaw + '\n' : '') + '＋ ' + extraRaw : r.voiceRaw,
    inputMethods: methods.join(','),
    extractedKeywords: result.extractedKeywords.join('、'),
    aiInterpretation: result.interpretation,
    candidates: JSON.stringify(result.candidates),
    updatedAt: new Date(),
  };
  withLock_(() => update_('Records', RECORD_COLS, r._row, patch));
  return { record: serializeRecord_(Object.assign({}, r, patch)), inference: result };
}

/** 正解を確定し、自動分類 */
function resolve_(body) {
  const user = requireUser_(body);
  const r = findOwnRecord_(user, body.recordId);
  const answer = String(body.answer || '').trim().slice(0, 200);
  if (!answer) throw new AppError('正解を入力してください。');
  const resolvedBy = body.resolvedBy === 'AI候補' ? 'AI候補' : '自力';

  let cls;
  try {
    cls = classify_(r.clues, answer, r.situation);
  } catch (err) {
    // 分類に失敗しても正解は保存する
    let fallback = 'その他';
    try {
      const hit = JSON.parse(r.candidates || '[]').find(c => c.name === answer);
      if (hit) fallback = hit.category;
    } catch (e) { /* noop */ }
    cls = { category: fallback, subCategory: '', tags: [], memoryHint: '' };
  }

  const resolvedAt = new Date();
  const created = r.createdAt instanceof Date ? r.createdAt : new Date(r.createdAt);
  const patch = {
    answer,
    resolvedBy,
    resolvedAt,
    resolveMinutes: Math.max(0, Math.round((resolvedAt - created) / 60000)),
    category: cls.category,
    subCategory: cls.subCategory,
    tags: cls.tags.join(','),
    memoryHint: cls.memoryHint,
    status: '解決',
    updatedAt: new Date(),
  };
  withLock_(() => update_('Records', RECORD_COLS, r._row, patch));
  return { record: serializeRecord_(Object.assign({}, r, patch)) };
}

/* =========================================================
 * 記録の修正
 * ======================================================= */

/** 数式として解釈されないよう、先頭が = + - @ の文字列は ' を付けて保存 */
function safeText_(v, max) {
  const t = String(v == null ? '' : v).trim().slice(0, max);
  return /^[=+\-@]/.test(t) ? "'" + t : t;
}

function normTags_(v) {
  const arr = Array.isArray(v) ? v : String(v || '').split(/[,、，\n]+/);
  const out = [];
  arr.map(t => String(t).replace(/^[#＃\s]+/, '').trim().slice(0, 30)).filter(Boolean).forEach(t => {
    if (out.indexOf(t) < 0) out.push(t);
  });
  return out.slice(0, 10);
}

/**
 * 記録の内容を修正する。fields に含まれる項目だけを更新。
 *   answer を空にすると「未解決」に戻し、分類・解決情報をクリア。
 *   未解決の記録に answer を入れると「解決」にし、カテゴリ未指定ならAIで分類。
 */
function updateRecord_(body) {
  const user = requireUser_(body);
  const f = Object.assign({}, body.fields || {});
  if (body.reopen) f.answer = ''; // 「解決を取り消して未解決に戻す」
  const has = k => Object.prototype.hasOwnProperty.call(f, k);

  // 分類AIはロックの外で呼ぶ（時間がかかるため）
  const before = findOwnRecord_(user, body.recordId);
  const clues = has('clues') ? String(f.clues || '').trim().slice(0, 3000) : String(before.clues || '');
  const situation = has('situation') ? (SITUATIONS.indexOf(f.situation) >= 0 ? f.situation : '') : String(before.situation || '');
  const answer = has('answer') ? String(f.answer || '').trim().slice(0, 200) : String(before.answer || '');
  const wantsResolved = !!answer;

  let autoCls = null;
  if (wantsResolved && has('category') && CATEGORIES.indexOf(f.category) < 0 && !before.category) {
    try { autoCls = classify_(clues, answer, situation); } catch (e) { autoCls = null; }
  }

  return withLock_(() => {
    const r = findOwnRecord_(user, body.recordId);
    const patch = {};

    if (has('clues')) {
      if (!clues && !String(r.imageFileIds || '')) throw new AppError('手がかりを空にはできません（画像のない記録です）。');
      patch.clues = safeText_(clues, 3000);
    }
    if (has('situation')) patch.situation = situation;
    if (has('visibility')) {
      const vis = f.visibility === VIS_PUBLIC ? VIS_PUBLIC : VIS_PRIVATE;
      patch.visibility = vis;
      if (vis === VIS_PUBLIC && r.visibility !== VIS_PUBLIC) patch.publishedAt = new Date();
    }
    if (has('occurredAt')) {
      const d = new Date(f.occurredAt);
      if (isNaN(d.getTime())) throw new AppError('発生日時の形式が正しくありません。');
      patch.occurredAt = d;
    }

    if (has('answer')) {
      if (!wantsResolved) {
        // 未解決に戻す
        Object.assign(patch, {
          answer: '', category: '', subCategory: '', tags: '', memoryHint: '',
          resolvedBy: '', resolvedAt: '', resolveMinutes: '', status: '未解決',
        });
      } else {
        patch.answer = safeText_(answer, 200);
        if (r.status !== '解決') {
          const resolvedAt = new Date();
          const created = r.createdAt instanceof Date ? r.createdAt : new Date(r.createdAt);
          patch.resolvedAt = resolvedAt;
          patch.resolveMinutes = isNaN(created.getTime()) ? '' : Math.max(0, Math.round((resolvedAt - created) / 60000));
          patch.resolvedBy = '自力';
          patch.status = '解決';
        }
      }
    }

    // 分類（解決状態のときだけ保存）
    if (wantsResolved) {
      if (has('category')) {
        if (CATEGORIES.indexOf(f.category) >= 0) patch.category = f.category;
        else if (autoCls) patch.category = autoCls.category;
        else if (!r.category) patch.category = 'その他';
      }
      if (has('subCategory')) patch.subCategory = safeText_(f.subCategory, 50);
      else if (autoCls) patch.subCategory = autoCls.subCategory;
      if (has('tags')) patch.tags = normTags_(f.tags).join(',');
      else if (autoCls) patch.tags = autoCls.tags.join(',');
      if (has('memoryHint')) patch.memoryHint = safeText_(f.memoryHint, 200);
      else if (autoCls) patch.memoryHint = autoCls.memoryHint;
      if (has('resolvedBy') && ['AI候補', '自力'].indexOf(f.resolvedBy) >= 0) patch.resolvedBy = f.resolvedBy;
      // 空欄で送られてきたがAI分類した項目は埋める
      if (autoCls) {
        if (!patch.subCategory) patch.subCategory = autoCls.subCategory;
        if (!patch.tags) patch.tags = autoCls.tags.join(',');
        if (!patch.memoryHint) patch.memoryHint = autoCls.memoryHint;
      }
    }

    patch.updatedAt = new Date();
    update_('Records', RECORD_COLS, r._row, patch);
    // 返却用は数式よけの ' を外す（シート上では表示されない接頭辞のため）
    const shown = {};
    Object.keys(patch).forEach(k => {
      const v = patch[k];
      shown[k] = typeof v === 'string' && /^'[=+\-@]/.test(v) ? v.slice(1) : v;
    });
    const record = serializeRecord_(Object.assign({}, r, shown));
    // patch は画面側で記録に上書きする値（記録全体を返す）
    return { record, patch: record, message: body.reopen ? '未解決に戻しました。' : '記録を修正しました。' };
  });
}

/** 修正画面の「AIに分類し直してもらう」：保存はせず、分類案だけ返す */
function reclassify_(body) {
  const user = requireUser_(body);
  const r = findOwnRecord_(user, body.recordId);
  const answer = String(body.answer != null ? body.answer : r.answer || '').trim().slice(0, 200);
  if (!answer) throw new AppError('分類するには正解を入力してください。');
  const clues = String(body.clues != null ? body.clues : r.clues || '').trim().slice(0, 3000);
  const situation = String(body.situation != null ? body.situation : r.situation || '');
  const cls = classify_(clues, answer, situation);
  return { classification: cls };
}

/* =========================================================
 * 記録：一覧・削除・画像取得
 * ======================================================= */

function serializeRecord_(r) {
  let candidates = [];
  try { candidates = JSON.parse(r.candidates || '[]'); } catch (e) { candidates = []; }
  const split = (v, sep) => String(v || '').split(sep).map(s => s.trim()).filter(Boolean);
  return {
    recordId: r.recordId,
    createdAt: toIso_(r.createdAt),
    occurredAt: toIso_(r.occurredAt),
    situation: r.situation || '',
    clues: r.clues || '',
    inputMethods: split(r.inputMethods, ','),
    imageUrls: split(r.imageUrls, '\n'),
    imageFileIds: split(r.imageFileIds, ','),
    imageKinds: split(r.imageKinds, ','),
    extractedKeywords: split(r.extractedKeywords, '、'),
    aiInterpretation: r.aiInterpretation || '',
    candidates,
    answer: r.answer || '',
    category: r.category || '',
    subCategory: r.subCategory || '',
    tags: split(r.tags, ','),
    memoryHint: r.memoryHint || '',
    resolvedBy: r.resolvedBy || '',
    resolvedAt: toIso_(r.resolvedAt),
    resolveMinutes: r.resolveMinutes === '' || r.resolveMinutes == null ? null : Number(r.resolveMinutes),
    status: r.status || '未解決',
    updatedAt: toIso_(r.updatedAt),
    visibility: r.visibility === VIS_PUBLIC ? VIS_PUBLIC : VIS_PRIVATE,
    publishedAt: toIso_(r.publishedAt),
  };
}

function listRecords_(body) {
  const user = requireUser_(body);
  const social = socialIndex_(user.userId);
  const list = readAll_('Records', RECORD_COLS)
    .filter(r => r.userId === user.userId)
    .map(r => Object.assign(serializeRecord_(r), { social: social[r.recordId] || emptySocial_() }))
    .sort((a, b) => (b.occurredAt || b.createdAt).localeCompare(a.occurredAt || a.createdAt));
  return { records: list };
}

function deleteRecord_(body) {
  const user = requireUser_(body);
  withLock_(() => {
    const r = findOwnRecord_(user, body.recordId);
    String(r.imageFileIds || '').split(',').filter(Boolean).forEach(id => {
      try { DriveApp.getFileById(id).setTrashed(true); } catch (e) { /* 既に無い */ }
    });
    sheet_('Records', RECORD_COLS).deleteRow(r._row);
    deleteRowsWhere_('Comments', COMMENT_COLS, c => c.recordId === r.recordId);
    deleteRowsWhere_('Reactions', REACTION_COLS, x => x.recordId === r.recordId);
  });
  return { message: '記録を削除しました。' };
}

/** 本人の記録、または公開中の記録に属する画像だけを base64 で返す（Driveは非公開のまま） */
function getImage_(body) {
  const user = requireUser_(body);
  const fileId = String(body.fileId || '');
  if (!fileId) throw new AppError('画像が見つかりません。');
  const ok = readAll_('Records', RECORD_COLS).some(r =>
    (r.userId === user.userId || r.visibility === VIS_PUBLIC) &&
    String(r.imageFileIds || '').split(',').indexOf(fileId) >= 0);
  if (!ok) throw new AppError('画像が見つかりません。');
  const blob = DriveApp.getFileById(fileId).getBlob();
  return { mimeType: blob.getContentType(), data: Utilities.base64Encode(blob.getBytes()) };
}

/* =========================================================
 * エクスポート
 * ======================================================= */

function exportToDrive_(body) {
  const user = requireUser_(body);
  const format = body.format === 'csv' ? 'csv' : 'md';
  const records = listRecords_(body).records;
  if (!records.length) throw new AppError('書き出す記録がまだありません。');
  const stamp = Utilities.formatDate(new Date(), TZ, 'yyyyMMdd_HHmm');
  const folder = userFolder_(user, 'exports');
  const d = iso => (iso ? Utilities.formatDate(new Date(iso), TZ, 'yyyy/MM/dd HH:mm') : '');

  let file;
  if (format === 'csv') {
    const header = ['発生日時', '状況', '入力手段', '手がかり', '正解', 'カテゴリ', 'サブカテゴリ', 'タグ', '解決方法', '解決までの分数', 'ステータス', '公開設定', '覚え方ヒント', '画像URL'];
    const esc = v => '"' + String(v == null ? '' : v).replace(/"/g, '""') + '"';
    const rows = records.map(r => [
      d(r.occurredAt), r.situation, r.inputMethods.join('/'), r.clues, r.answer, r.category, r.subCategory,
      r.tags.join('/'), r.resolvedBy, r.resolveMinutes == null ? '' : r.resolveMinutes, r.status, r.visibility, r.memoryHint, r.imageUrls.join(' '),
    ].map(esc).join(','));
    const csv = '\uFEFF' + [header.map(esc).join(',')].concat(rows).join('\r\n');
    file = folder.createFile(Utilities.newBlob(csv, 'text/csv', `もやもやメモ_${stamp}.csv`));
  } else {
    const lines = [`# ${APP_NAME} 記録（${user.displayName}）`, '', `書き出し: ${d(new Date().toISOString())}　件数: ${records.length}`, ''];
    records.forEach(r => {
      lines.push(`## ${d(r.occurredAt)}　${r.answer || '（未解決）'}`);
      lines.push('');
      lines.push(`- ステータス: ${r.status}${r.resolvedBy ? `（${r.resolvedBy}）` : ''}`);
      lines.push(`- 公開設定: ${r.visibility}${r.social && r.visibility === VIS_PUBLIC ? `（あるある ${r.social.aruaru}・いいね ${r.social.like}・コメント ${r.social.comments}）` : ''}`);
      if (r.category) lines.push(`- カテゴリ: ${r.category}${r.subCategory ? ' / ' + r.subCategory : ''}`);
      if (r.tags.length) lines.push(`- タグ: ${r.tags.map(t => '#' + t.replace(/\s/g, '_')).join(' ')}`);
      lines.push(`- 状況: ${r.situation || '-'}`);
      lines.push(`- 入力手段: ${r.inputMethods.join('、') || '-'}`);
      if (r.resolveMinutes != null) lines.push(`- 解決まで: ${r.resolveMinutes}分`);
      lines.push('');
      lines.push('> ' + (r.clues || '（画像のみ）').replace(/\n/g, '\n> '));
      lines.push('');
      if (r.memoryHint) lines.push(`覚え方: ${r.memoryHint}`, '');
      r.imageUrls.forEach((u, i) => lines.push(`- [画像${i + 1}](${u})`));
      lines.push('');
    });
    file = folder.createFile(Utilities.newBlob(lines.join('\n'), 'text/markdown', `もやもやメモ_${stamp}.md`));
  }
  return { url: file.getUrl(), name: file.getName() };
}

/* =========================================================
 * みんなのもやもや：公開設定・一覧・あるある/いいね・コメント
 * ======================================================= */

function emptySocial_() {
  return { aruaru: 0, like: 0, comments: 0, mine: { aruaru: false, like: false } };
}

/** 記録IDごとの反応数・コメント数（userId の反応は mine に） */
function socialIndex_(userId) {
  const idx = {};
  const get = id => idx[id] || (idx[id] = emptySocial_());
  readAll_('Reactions', REACTION_COLS).forEach(x => {
    if (!REACTIONS[x.type]) return;
    const s = get(x.recordId);
    s[x.type]++;
    if (x.userId === userId) s.mine[x.type] = true;
  });
  readAll_('Comments', COMMENT_COLS).forEach(c => { get(c.recordId).comments++; });
  return idx;
}

function userNames_() {
  const m = {};
  readAll_('Users', USER_COLS).forEach(u => { m[u.userId] = String(u.displayName || ''); });
  return m;
}

function deleteRowsWhere_(name, cols, pred) {
  const rows = readAll_(name, cols).filter(pred).map(x => x._row).sort((a, b) => b - a);
  if (!rows.length) return;
  const sh = sheet_(name, cols);
  rows.forEach(n => sh.deleteRow(n));
}

/** 本人の記録か、公開中の記録だけ読める */
function findReadableRecord_(user, recordId) {
  const r = readAll_('Records', RECORD_COLS).find(x => x.recordId === recordId);
  if (!r || (r.userId !== user.userId && r.visibility !== VIS_PUBLIC)) {
    throw new AppError('この記録は見つからないか、非公開になりました。', 'GONE');
  }
  return r;
}

/** 公開用の記録（音声原文・Drive URL・AI候補などは出さない） */
function publicRecord_(r, user, social, names) {
  const s = serializeRecord_(r);
  const resolved = s.status === '解決';
  return {
    recordId: s.recordId,
    author: names[r.userId] || '退会したユーザー',
    isMine: r.userId === user.userId,
    occurredAt: s.occurredAt,
    publishedAt: s.publishedAt || s.updatedAt || s.createdAt,
    situation: s.situation,
    clues: s.clues,
    aiInterpretation: resolved ? '' : s.aiInterpretation,
    answer: s.answer,
    category: s.category,
    subCategory: s.subCategory,
    tags: s.tags,
    memoryHint: s.memoryHint,
    status: s.status,
    resolvedBy: s.resolvedBy,
    resolveMinutes: s.resolveMinutes,
    inputMethods: s.inputMethods,
    imageFileIds: s.imageFileIds,
    imageKinds: s.imageKinds,
    visibility: s.visibility,
    social: social[r.recordId] || emptySocial_(),
  };
}

/** 記録ごとの公開・非公開の切り替え */
function setVisibility_(body) {
  const user = requireUser_(body);
  const vis = body.visibility === VIS_PUBLIC ? VIS_PUBLIC : VIS_PRIVATE;
  return withLock_(() => {
    const r = findOwnRecord_(user, body.recordId);
    const patch = { visibility: vis, updatedAt: new Date() };
    if (vis === VIS_PUBLIC && r.visibility !== VIS_PUBLIC) patch.publishedAt = new Date();
    update_('Records', RECORD_COLS, r._row, patch);
    return {
      record: serializeRecord_(Object.assign({}, r, patch)),
      message: vis === VIS_PUBLIC ? 'みんなに公開しました。' : '非公開にしました。',
    };
  });
}

/** 公開中の記録一覧（新しく公開された順） */
function listPublic_(body) {
  const user = requireUser_(body);
  const filter = String(body.filter || 'all');
  const q = String(body.q || '').trim().toLowerCase().slice(0, 50);
  const offset = Math.max(0, Number(body.offset) || 0);
  const limit = Math.min(50, Math.max(1, Number(body.limit) || PUBLIC_PAGE));

  let list = readAll_('Records', RECORD_COLS).filter(r => r.visibility === VIS_PUBLIC);
  if (filter === 'unresolved') list = list.filter(r => r.status !== '解決');
  else if (filter === 'resolved') list = list.filter(r => r.status === '解決');
  else if (filter === 'mine') list = list.filter(r => r.userId === user.userId);
  if (q) {
    list = list.filter(r => [r.clues, r.answer, r.category, r.subCategory, r.tags, r.situation]
      .join(' ').toLowerCase().indexOf(q) >= 0);
  }
  const time = r => {
    const v = r.publishedAt || r.updatedAt || r.createdAt;
    const t = (v instanceof Date ? v : new Date(v)).getTime();
    return isNaN(t) ? 0 : t;
  };
  list.sort((a, b) => time(b) - time(a));

  const total = list.length;
  const page = list.slice(offset, offset + limit);
  const social = socialIndex_(user.userId);
  const names = userNames_();
  return {
    items: page.map(r => publicRecord_(r, user, social, names)),
    total,
    nextOffset: offset + page.length,
    hasMore: offset + page.length < total,
  };
}

/** コメント一覧と反応数（コメントした人・記録した人は修正・削除できる） */
function commentsPayload_(user, r) {
  const names = userNames_();
  const t = v => { const d = v instanceof Date ? v : new Date(v); return isNaN(d.getTime()) ? 0 : d.getTime(); };
  const comments = readAll_('Comments', COMMENT_COLS)
    .filter(c => c.recordId === r.recordId)
    .sort((a, b) => t(a.createdAt) - t(b.createdAt))
    .map(c => ({
      commentId: c.commentId,
      author: names[c.userId] || '退会したユーザー',
      isMine: c.userId === user.userId,
      isRecordOwner: c.userId === r.userId,
      canEdit: c.userId === user.userId || r.userId === user.userId,
      text: String(c.text || ''),
      createdAt: toIso_(c.createdAt),
      updatedAt: toIso_(c.updatedAt),
      edited: t(c.updatedAt) - t(c.createdAt) > 1000,
      editedByOwner: !!c.editedBy && c.editedBy !== c.userId,
    }));
  const social = socialIndex_(user.userId)[r.recordId] || emptySocial_();
  return { recordId: r.recordId, comments, social, isPublic: r.visibility === VIS_PUBLIC, isMine: r.userId === user.userId };
}

function getComments_(body) {
  const user = requireUser_(body);
  return commentsPayload_(user, findReadableRecord_(user, body.recordId));
}

function cleanComment_(v) {
  const text = String(v || '').replace(/\r\n?/g, '\n').trim();
  if (!text) throw new AppError('コメントを入力してください。');
  if (text.length > COMMENT_MAX) throw new AppError(`コメントは${COMMENT_MAX}文字以内にしてください。`);
  return text;
}

function addComment_(body) {
  const user = requireUser_(body);
  const text = cleanComment_(body.text);
  return withLock_(() => {
    const r = findReadableRecord_(user, body.recordId);
    if (r.visibility !== VIS_PUBLIC) throw new AppError('非公開の記録にはコメントできません。');
    const now = new Date();
    append_('Comments', COMMENT_COLS, {
      commentId: 'C' + Utilities.formatDate(now, TZ, 'yyMMddHHmmss') + Utilities.getUuid().slice(0, 6),
      recordId: r.recordId,
      userId: user.userId,
      text: safeText_(text, COMMENT_MAX),
      createdAt: now,
      updatedAt: now,
      editedBy: '',
    });
    return Object.assign(commentsPayload_(user, r), { message: 'コメントしました。' });
  });
}

/** 修正・削除の権限確認：コメントした本人か、コメントされた記録の持ち主 */
function findEditableComment_(user, commentId) {
  const c = readAll_('Comments', COMMENT_COLS).find(x => x.commentId === commentId);
  if (!c) throw new AppError('コメントが見つかりません（すでに削除された可能性があります）。');
  const r = readAll_('Records', RECORD_COLS).find(x => x.recordId === c.recordId) || null;
  const allowed = c.userId === user.userId || (r && r.userId === user.userId);
  if (!allowed) throw new AppError('このコメントを修正・削除する権限がありません。');
  return { c, r };
}

function afterCommentChange_(user, r, message) {
  try {
    if (r) return Object.assign(commentsPayload_(user, findReadableRecord_(user, r.recordId)), { message });
  } catch (e) { /* 非公開になった記録など */ }
  return { message, comments: null };
}

function editComment_(body) {
  const user = requireUser_(body);
  const text = cleanComment_(body.text);
  return withLock_(() => {
    const { c, r } = findEditableComment_(user, String(body.commentId || ''));
    update_('Comments', COMMENT_COLS, c._row, {
      text: safeText_(text, COMMENT_MAX),
      updatedAt: new Date(),
      editedBy: user.userId,
    });
    return afterCommentChange_(user, r, 'コメントを修正しました。');
  });
}

function deleteComment_(body) {
  const user = requireUser_(body);
  return withLock_(() => {
    const { c, r } = findEditableComment_(user, String(body.commentId || ''));
    sheet_('Comments', COMMENT_COLS).deleteRow(c._row);
    return afterCommentChange_(user, r, 'コメントを削除しました。');
  });
}

/** 「それ、あるある」「いいね」を付ける／外す */
function toggleReaction_(body) {
  const user = requireUser_(body);
  const type = Object.prototype.hasOwnProperty.call(REACTIONS, body.type) ? body.type : '';
  if (!type) throw new AppError('不明な反応です。');
  return withLock_(() => {
    const r = findReadableRecord_(user, body.recordId);
    if (r.visibility !== VIS_PUBLIC) throw new AppError('非公開の記録には付けられません。');
    if (r.userId === user.userId) throw new AppError('自分の記録には付けられません。');
    const ex = readAll_('Reactions', REACTION_COLS)
      .find(x => x.recordId === r.recordId && x.userId === user.userId && x.type === type);
    if (ex) sheet_('Reactions', REACTION_COLS).deleteRow(ex._row);
    else append_('Reactions', REACTION_COLS, { recordId: r.recordId, userId: user.userId, type, createdAt: new Date() });
    return { recordId: r.recordId, social: socialIndex_(user.userId)[r.recordId] || emptySocial_() };
  });
}
