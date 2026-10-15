/**
 * 認証・セッション管理
 *
 * 口座情報や金額を扱うエンドポイントを、PIN（指導者）または事務局パスワード
 * （事務局）で発行した短命トークンの保持者だけに限定する。
 *
 * 【設計】
 * - トークンは「セッション」シートに保存する。スクリプトプロパティだと
 *   同時書き込みで失われる恐れがあり、監査用に発行履歴も残したいため。
 * - トークンは URL に乗せない。GASの実行ログやブラウザ履歴に残るため、
 *   認証が必要なアクションはすべて POST（本文にトークン）で受ける。
 * - 指導者トークンは「自分のデータ」にしかアクセスできない。
 *   事務局トークンは全データにアクセスできる。
 *
 * 【事前準備】プロジェクトの設定 → スクリプト プロパティ
 *   ADMIN_PASSWORD : 事務局画面の共通パスワード
 */

// トークンの有効期間（ミリ秒）。指導者・事務局ともに8時間。
const TOKEN_TTL_MS = 8 * 60 * 60 * 1000;

const SESSION_SHEET_NAME = 'セッション';

const ROLE = {
  INSTRUCTOR: 'instructor',
  ADMIN: 'admin',
};

// ========== セッションシート ==========

function getOrCreateSessionSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(SESSION_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SESSION_SHEET_NAME);
    sheet.appendRow(['トークン', '種別', '氏名', '発行日時', '有効期限']);
    sheet.setFrozenRows(1);
    // 運用で開く必要のないシートなので隠す
    sheet.hideSheet();
  }
  return sheet;
}

const SESSION_COL = {
  TOKEN: 0,
  ROLE: 1,
  NAME: 2,
  ISSUED_AT: 3,
  EXPIRES_AT: 4,
};

// ========== トークン発行 ==========

/**
 * 推測困難なトークンを生成する。
 * UUIDを2つ連結し、ハイフンを除いた64文字の16進文字列にする。
 */
function generateToken() {
  return (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
}

/**
 * トークンを発行してセッションシートに記録する。
 * 発行のついでに期限切れ行を掃除する。
 */
function issueToken(role, name) {
  return withReportLock(function () {
    const sheet = getOrCreateSessionSheet();
    purgeExpiredSessions(sheet);

    const token = generateToken();
    const now = new Date();
    const expiresAt = new Date(now.getTime() + TOKEN_TTL_MS);

    sheet.appendRow([token, role, name || '', now, expiresAt]);
    SpreadsheetApp.flush();

    Logger.log('[issueToken] 発行: 種別=%s, 氏名=%s, 期限=%s', role, name || '(なし)', expiresAt);
    return { token: token, expiresAt: expiresAt.getTime() };
  });
}

/**
 * 期限切れのセッション行を削除する（呼び出し元でロックを取っていること）
 */
function purgeExpiredSessions(sheet) {
  const data = sheet.getDataRange().getValues();
  const now = new Date().getTime();
  let deleted = 0;

  // 後ろから削除してインデックスずれを防ぐ
  for (let i = data.length - 1; i >= 1; i--) {
    const exp = data[i][SESSION_COL.EXPIRES_AT];
    const expMs = exp instanceof Date ? exp.getTime() : new Date(exp || 0).getTime();
    if (!expMs || expMs < now) {
      sheet.deleteRow(i + 1);
      deleted++;
    }
  }
  if (deleted > 0) Logger.log('[purgeExpiredSessions] 期限切れ %s 件を削除', deleted);
}

// ========== トークン検証 ==========

/**
 * トークンを検証し、{ role, name } を返す。無効なら null。
 */
function verifyToken(token) {
  if (!token) return null;

  const sheet = getOrCreateSessionSheet();
  const data = sheet.getDataRange().getValues();
  const target = String(token).trim();
  const now = new Date().getTime();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][SESSION_COL.TOKEN]).trim() !== target) continue;

    const exp = data[i][SESSION_COL.EXPIRES_AT];
    const expMs = exp instanceof Date ? exp.getTime() : new Date(exp || 0).getTime();
    if (!expMs || expMs < now) {
      Logger.log('[verifyToken] 期限切れトークン');
      return null;
    }
    return {
      role: String(data[i][SESSION_COL.ROLE]).trim(),
      name: String(data[i][SESSION_COL.NAME]).trim(),
    };
  }
  return null;
}

/**
 * 認証が必要なアクションの入口で呼ぶ。
 *
 * requiredRole:
 *   ROLE.ADMIN      → 事務局トークンのみ許可
 *   ROLE.INSTRUCTOR → 指導者トークンまたは事務局トークンを許可。
 *                     指導者トークンの場合、body.instructorName が本人と
 *                     一致しなければ拒否する（他人のデータを見せない）。
 *
 * 認証できない場合は AUTH_REQUIRED を含むエラーを投げる。
 * フロントはこの文字列を見て再ログインを促す。
 */
function requireAuth(body, requiredRole) {
  const auth = verifyToken(body && body.token);
  if (!auth) {
    throw new Error('AUTH_REQUIRED:再ログインが必要です');
  }

  if (auth.role === ROLE.ADMIN) return auth;   // 事務局は全て許可

  if (requiredRole === ROLE.ADMIN) {
    throw new Error('AUTH_FORBIDDEN:この操作には事務局の権限が必要です');
  }

  // 指導者トークン: 自分のデータのみ
  const requested = String((body && body.instructorName) || '').trim();
  if (requested && requested !== auth.name) {
    Logger.log('[requireAuth] 他人のデータへのアクセスを拒否: token=%s, 要求=%s', auth.name, requested);
    throw new Error('AUTH_FORBIDDEN:他の指導者のデータは参照できません');
  }
  return auth;
}

// ========== 事務局ログイン ==========

/**
 * 事務局パスワードを照合してトークンを発行する。
 * body: { password }
 */
function adminLogin(body) {
  const expected = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  if (!expected) {
    Logger.log('[adminLogin] ADMIN_PASSWORD が未設定');
    return {
      success: false,
      error: '事務局パスワードが未設定です。スクリプトプロパティに ADMIN_PASSWORD を設定してください',
    };
  }

  const given = String((body && body.password) || '');
  if (!given) {
    return { success: false, error: 'パスワードを入力してください' };
  }
  if (!constantTimeEquals(given, String(expected))) {
    Logger.log('[adminLogin] パスワード不一致');
    return { success: false, error: 'パスワードが正しくありません' };
  }

  const issued = issueToken(ROLE.ADMIN, '事務局');
  return { success: true, token: issued.token, expiresAt: issued.expiresAt, role: ROLE.ADMIN };
}

/**
 * 文字列比較にかかる時間を入力内容で変えないようにする。
 * 先頭から何文字一致したかで総当たりを効率化されるのを防ぐ。
 */
function constantTimeEquals(a, b) {
  const sa = String(a);
  const sb = String(b);
  let diff = sa.length ^ sb.length;
  const len = Math.max(sa.length, sb.length);
  for (let i = 0; i < len; i++) {
    diff |= (sa.charCodeAt(i) || 0) ^ (sb.charCodeAt(i) || 0);
  }
  return diff === 0;
}

// ========== ログアウト ==========

/**
 * トークンを無効化する。
 * body: { token }
 */
function logout(body) {
  const token = String((body && body.token) || '').trim();
  if (!token) return { success: true };

  return withReportLock(function () {
    const sheet = getOrCreateSessionSheet();
    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][SESSION_COL.TOKEN]).trim() === token) {
        sheet.deleteRow(i + 1);
      }
    }
    SpreadsheetApp.flush();
    return { success: true };
  });
}

// ========== 保守 ==========

/**
 * 全セッションを失効させる（パスワード変更時や漏洩時に手動実行する）
 */
function revokeAllSessions() {
  return withReportLock(function () {
    const sheet = getOrCreateSessionSheet();
    const last = sheet.getLastRow();
    if (last > 1) sheet.deleteRows(2, last - 1);
    SpreadsheetApp.flush();
    Logger.log('[revokeAllSessions] 全セッションを失効させました（%s 件）', Math.max(last - 1, 0));
    return { success: true, revoked: Math.max(last - 1, 0) };
  });
}

/**
 * 事務局パスワードを設定する（この関数の引数を書き換えて1回だけ実行する）
 * 例: setAdminPassword('ここに新しいパスワード');
 *
 * ※パスワードをコードに残さないため、実行後は引数を消しておくこと。
 * ※変更後は revokeAllSessions() を実行して既存セッションを失効させる。
 */
function setAdminPassword(password) {
  if (!password || String(password).length < 8) {
    Logger.log('パスワードは8文字以上にしてください。例: setAdminPassword(\'...\')');
    return { success: false };
  }
  PropertiesService.getScriptProperties().setProperty('ADMIN_PASSWORD', String(password));
  Logger.log('事務局パスワードを設定しました。revokeAllSessions() で既存セッションを失効させてください。');
  return { success: true };
}
