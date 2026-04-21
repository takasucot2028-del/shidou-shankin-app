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
      case 'getReport':
        result = getReport(e.parameter);
        break;
      case 'getInstructors':
        result = getInstructors();
        break;
      case 'getDashboard':
        result = getDashboard(e.parameter);
        break;
      case 'exportSheet':
        result = exportSheet(e.parameter);
        break;
      default:
        result = { success: false, error: 'Unknown action: ' + action };
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
      case 'saveReport':
        result = saveReport(body);
        break;
      case 'submitReport':
        result = submitReport(body);
        break;
      case 'updateMaster':
        result = updateMaster(body);
        break;
      case 'calcFee':
        result = calcFee(body);
        break;
      case 'calcAllFees':
        result = calcAllFees(body);
        break;
      case 'updateFee':
        result = updateFee(body);
        break;
      case 'generateTransferSheet':
        result = generateTransferSheet(body);
        break;
      case 'generateClubSummarySheet':
        result = generateClubSummarySheet(body);
        break;
      case 'sendPaySlipEmails':
        result = sendPaySlipEmails(body);
        break;
      case 'updateReport':
        result = updateReport(body);
        break;
      default:
        result = { success: false, error: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { success: false, error: err.message };
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
      '謝金計算時間(h)', '交通手段', '行先', '旅費金額', '備考',
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
