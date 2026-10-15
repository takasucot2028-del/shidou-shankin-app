/**
 * GAS メインファイル
 * Web App として doGet / doPost でリクエストを受け取り、各モジュールに委譲する
 */

// スプレッドシートID（GASのスクリプトプロパティに設定してください）
// PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', 'YOUR_ID');
function getSpreadsheetId() {
  return PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
}

// 事務局メールアドレス一覧（カンマ区切りでスクリプトプロパティに設定）
function getAdminEmails() {
  const raw = PropertiesService.getScriptProperties().getProperty('ADMIN_EMAILS') || '';
  return raw.split(',').map(e => e.trim()).filter(Boolean);
}

// ========== エントリーポイント ==========

function doGet(e) {
  const action = e.parameter.action;
  let result;

  try {
    switch (action) {
      // ===== 認証不要（個人情報を含まないもののみ） =====
      case 'getInstructorsPublic':
        result = getInstructorsPublic();
        break;

      default:
        // 口座情報・金額・月報を返すアクションはGETでは受け付けない。
        // トークンをURLに乗せるとGASの実行ログやブラウザ履歴に残るため、
        // 認証が必要なものはすべてPOST（本文にトークン）で受ける。
        // POSTがリダイレクトの際にGETへ変換され、本文（アクション・トークン）が
        // 失われるとここに来る。この場合アクションは一切実行されていないため、
        // フロントは getNotAllowed を見て安全に再送できる。
        result = {
          success: false,
          getNotAllowed: true,
          error: 'このアクションはGETでは利用できません。POSTでトークンを添えてください: ' + action,
        };
    }
  } catch (err) {
    result = { success: false, error: err.message };
  }

  return buildResponse(result);
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return buildResponse({ success: false, error: 'Invalid JSON body' });
  }

  const action = body.action;
  let result;

  try {
    switch (action) {
      // ===== 認証不要（ログイン処理そのもの） =====
      case 'checkPin':
        result = checkPin(body);
        break;
      case 'adminLogin':
        result = adminLogin(body);
        break;
      case 'logout':
        result = logout(body);
        break;
      case 'getInstructorsPublic':
        result = getInstructorsPublic();
        break;

      // ===== 指導者本人または事務局 =====
      // requireAuth が body.instructorName と本人の一致を検証する
      case 'getReport':
        requireAuth(body, ROLE.INSTRUCTOR);
        result = getReport(body);
        break;
      case 'getPaySlip':
        requireAuth(body, ROLE.INSTRUCTOR);
        result = getPaySlip(body);
        break;
      case 'saveReport':
        requireAuth(body, ROLE.INSTRUCTOR);
        result = saveReport(body);
        break;
      case 'submitReport':
        requireAuth(body, ROLE.INSTRUCTOR);
        result = submitReport(body);
        break;

      // ===== 事務局のみ =====
      case 'getInstructors':
        requireAuth(body, ROLE.ADMIN);
        result = getInstructors();
        break;
      case 'getDashboard':
        requireAuth(body, ROLE.ADMIN);
        result = getDashboard(body);
        break;
      case 'exportSheet':
        requireAuth(body, ROLE.ADMIN);
        result = exportSheet(body);
        break;
      case 'updateMaster':
        requireAuth(body, ROLE.ADMIN);
        result = updateMaster(body);
        break;
      case 'updateReport':
        requireAuth(body, ROLE.ADMIN);
        result = updateReport(body);
        break;
      case 'calcFee':
        requireAuth(body, ROLE.ADMIN);
        result = calcFee(body);
        break;
      case 'calcAllFees':
        requireAuth(body, ROLE.ADMIN);
        result = calcAllFees(body);
        break;
      case 'updateFee':
        requireAuth(body, ROLE.ADMIN);
        result = updateFee(body);
        break;
      case 'generateTransferSheet':
        requireAuth(body, ROLE.ADMIN);
        result = generateTransferSheet(body);
        break;
      case 'generateClubSummarySheet':
        requireAuth(body, ROLE.ADMIN);
        result = generateClubSummarySheet(body);
        break;
      case 'generateCategorySummarySheet':
        requireAuth(body, ROLE.ADMIN);
        result = generateCategorySummarySheet(body);
        break;
      case 'sendPaySlipEmails':
        requireAuth(body, ROLE.ADMIN);
        result = sendPaySlipEmails(body);
        break;

      default:
        result = { success: false, error: 'Unknown action: ' + action };
    }
  } catch (err) {
    // AUTH_REQUIRED / AUTH_FORBIDDEN はフロントが再ログイン判定に使う
    const msg = String(err.message || err);
    if (msg.indexOf('AUTH_REQUIRED:') === 0) {
      result = { success: false, authRequired: true, error: msg.replace('AUTH_REQUIRED:', '') };
    } else if (msg.indexOf('AUTH_FORBIDDEN:') === 0) {
      result = { success: false, forbidden: true, error: msg.replace('AUTH_FORBIDDEN:', '') };
    } else {
      result = { success: false, error: msg };
    }
  }

  return buildResponse(result);
}

// ========== レスポンスヘルパー ==========

function buildResponse(data) {
  const output = ContentService.createTextOutput(JSON.stringify(data));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

// ========== スプレッドシート取得ヘルパー ==========

function getSpreadsheet() {
  const id = getSpreadsheetId();
  if (!id) throw new Error('SPREADSHEET_ID が設定されていません');
  return SpreadsheetApp.openById(id);
}

function getSheet(sheetName) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('シートが見つかりません: ' + sheetName);
  return sheet;
}

// 年度ごとの月報シート名を返す
function getReportSheetName(year) {
  return '月報データ_' + year;
}

// 月報シートを取得（なければ作成）
function getOrCreateReportSheet(year) {
  const ss = getSpreadsheet();
  const name = getReportSheetName(year);
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    const headers = [
      '提出ID', '指導者氏名', 'クラブ名', '対象年', '対象月', '日付',
      '区分', '時給区分', '開始時刻', '終了時刻', '指導時間(h)',
      '謝金計算時間(h)', '交通手段', '場所', '旅費金額', '備考',
      'ステータス', '提出日時', '最終更新日時'
    ];
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ========== UUID生成 ==========

function generateUUID() {
  return Utilities.getUuid();
}
