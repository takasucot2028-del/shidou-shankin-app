/**
 * 月報 CRUD 処理
 * シート列定義（1-indexed に対応した0-indexed定数）
 */
const COL = {
  SUBMIT_ID: 0,
  INSTRUCTOR_NAME: 1,
  CLUB_NAME: 2,
  YEAR: 3,
  MONTH: 4,
  DATE: 5,
  CATEGORY: 6,
  RATE_TYPE: 7,
  START_TIME: 8,
  END_TIME: 9,
  INSTRUCTION_HOURS: 10,
  CALC_HOURS: 11,
  TRANSPORT: 12,
  DESTINATION: 13,
  TRAVEL_AMOUNT: 14,
  NOTE: 15,
  STATUS: 16,
  SUBMITTED_AT: 17,
  UPDATED_AT: 18,
};

// ========== 月報取得 ==========

/**
 * 月報データ取得
 * params: { instructorName, year, month, submitId }
 */
function getReport(params) {
  const year = params.year ? parseInt(params.year) : new Date().getFullYear();
  const sheet = getOrCreateReportSheet(year);
  const data = sheet.getDataRange().getValues();

  // ヘッダー行を除く
  const rows = data.slice(1).filter(row => {
    if (params.submitId && row[COL.SUBMIT_ID] !== params.submitId) return false;
    if (params.instructorName && row[COL.INSTRUCTOR_NAME] !== params.instructorName) return false;
    if (params.month && String(row[COL.MONTH]) !== String(params.month)) return false;
    return true;
  });

  return {
    success: true,
    data: rows.map(rowToReport),
  };
}

function rowToReport(row) {
  return {
    submitId: row[COL.SUBMIT_ID],
    instructorName: row[COL.INSTRUCTOR_NAME],
    clubName: row[COL.CLUB_NAME],
    year: row[COL.YEAR],
    month: row[COL.MONTH],
    date: row[COL.DATE],
    category: row[COL.CATEGORY],
    rateType: row[COL.RATE_TYPE],
    startTime: row[COL.START_TIME],
    endTime: row[COL.END_TIME],
    instructionHours: row[COL.INSTRUCTION_HOURS],
    calcHours: row[COL.CALC_HOURS],
    transport: row[COL.TRANSPORT],
    destination: row[COL.DESTINATION],
    travelAmount: row[COL.TRAVEL_AMOUNT],
    note: row[COL.NOTE],
    status: row[COL.STATUS],
    submittedAt: row[COL.SUBMITTED_AT],
    updatedAt: row[COL.UPDATED_AT],
  };
}

// ========== 月報一時保存 ==========

/**
 * 月報一時保存（下書き）
 * body: { instructorName, clubName, year, month, rows: [...] }
 * 同じ指導者・年月の既存下書き行を削除して再挿入する
 */
function saveReport(body) {
  validateReportBody(body);

  const year = parseInt(body.year);
  const month = parseInt(body.month);
  const sheet = getOrCreateReportSheet(year);

  // 既存の下書き行を削除（提出済みは残す）
  deleteReportRows(sheet, body.instructorName, year, month, '下書き');

  const now = new Date();
  // submitId は最初の保存時に生成し、以降は同じものを使う
  const submitId = body.submitId || generateUUID();

  const newRows = (body.rows || []).map(r => buildRow(r, {
    submitId,
    instructorName: body.instructorName,
    clubName: body.clubName,
    year,
    month,
    status: '下書き',
    submittedAt: '',
    updatedAt: now,
  }));

  if (newRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length)
      .setValues(newRows);
  }

  return { success: true, submitId };
}

// ========== 月報提出 ==========

/**
 * 月報提出（ステータスを「提出済」に変更）
 * body: { instructorName, clubName, year, month, rows: [...], submitId }
 * 既存の下書きを削除して提出済み行を挿入する
 */
function submitReport(body) {
  validateReportBody(body);

  const year = parseInt(body.year);
  const month = parseInt(body.month);
  const sheet = getOrCreateReportSheet(year);

  // 既に提出済みなら二重提出エラー
  if (hasSubmittedReport(sheet, body.instructorName, year, month)) {
    return { success: false, error: 'すでに提出済みです' };
  }

  // 下書き行を削除
  deleteReportRows(sheet, body.instructorName, year, month, '下書き');

  const now = new Date();
  const submitId = body.submitId || generateUUID();

  const newRows = (body.rows || []).map(r => buildRow(r, {
    submitId,
    instructorName: body.instructorName,
    clubName: body.clubName,
    year,
    month,
    status: '提出済',
    submittedAt: now,
    updatedAt: now,
  }));

  if (newRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length)
      .setValues(newRows);
  }

  // setValuesのバッファをシートにコミットしてから謝金計算を実行
  SpreadsheetApp.flush();
  calcFee({ instructorName: body.instructorName, year, month });

  // 事務局へ提出通知メール
  notifyAdminOnSubmit(body.instructorName, body.clubName, year, month, now);

  return { success: true, submitId };
}

// ========== ダッシュボードデータ取得 ==========

/**
 * 指定年月の提出状況一覧を返す
 * params: { year, month }
 */
function getDashboard(params) {
  const year = parseInt(params.year) || new Date().getFullYear();
  const month = parseInt(params.month) || new Date().getMonth() + 1;

  const sheet = getOrCreateReportSheet(year);
  const data = sheet.getDataRange().getValues().slice(1);

  // 指導者マスタ全員を取得
  const masterSheet = getSheet('指導者マスタ');
  const masterData = masterSheet.getDataRange().getValues().slice(1);
  const allInstructors = masterData.map(r => ({
    name: r[1],
    clubName: r[2],
    email: r[10],
  }));

  // 対象月の提出済み指導者を集計
  const submittedMap = {};
  data.forEach(row => {
    if (parseInt(row[COL.YEAR]) === year &&
        parseInt(row[COL.MONTH]) === month &&
        row[COL.STATUS] === '提出済') {
      submittedMap[row[COL.INSTRUCTOR_NAME]] = {
        clubName: row[COL.CLUB_NAME],
        submittedAt: row[COL.SUBMITTED_AT],
        status: '提出済',
      };
    }
  });

  const statusList = allInstructors.map(inst => ({
    instructorName: inst.name,
    clubName: inst.clubName,
    status: submittedMap[inst.name] ? '提出済' : '未提出',
    submittedAt: submittedMap[inst.name] ? submittedMap[inst.name].submittedAt : '',
  }));

  const submittedCount = statusList.filter(s => s.status === '提出済').length;
  const unsubmittedList = statusList.filter(s => s.status === '未提出');

  return {
    success: true,
    year,
    month,
    totalCount: allInstructors.length,
    submittedCount,
    unsubmittedList,
    statusList,
  };
}

// ========== ヘルパー ==========

function validateReportBody(body) {
  if (!body.instructorName) throw new Error('指導者氏名が必要です');
  if (!body.year) throw new Error('対象年が必要です');
  if (!body.month) throw new Error('対象月が必要です');
}

function buildRow(r, meta) {
  return [
    meta.submitId,
    meta.instructorName,
    meta.clubName,
    meta.year,
    meta.month,
    r.date || '',
    r.category || '',
    r.rateType || '',
    r.startTime || '',
    r.endTime || '',
    r.instructionHours || 0,
    r.calcHours || 0,
    r.transport || '',
    r.destination || '',
    r.travelAmount || 0,
    r.note || '',
    meta.status,
    meta.submittedAt,
    meta.updatedAt,
  ];
}

function deleteReportRows(sheet, instructorName, year, month, status) {
  const data = sheet.getDataRange().getValues();
  // 後ろから削除してインデックスずれを防ぐ
  for (let i = data.length - 1; i >= 1; i--) {
    const row = data[i];
    if (row[COL.INSTRUCTOR_NAME] === instructorName &&
        parseInt(row[COL.YEAR]) === year &&
        parseInt(row[COL.MONTH]) === month &&
        row[COL.STATUS] === status) {
      sheet.deleteRow(i + 1); // シートは1-indexed
    }
  }
}

function hasSubmittedReport(sheet, instructorName, year, month) {
  const data = sheet.getDataRange().getValues().slice(1);
  return data.some(row =>
    row[COL.INSTRUCTOR_NAME] === instructorName &&
    parseInt(row[COL.YEAR]) === year &&
    parseInt(row[COL.MONTH]) === month &&
    row[COL.STATUS] === '提出済'
  );
}

// ========== エクスポート ==========

/**
 * 謝金計算結果シートの内容をJSONで返す
 * params: { year, month }
 */
function exportSheet(params) {
  const calcSheet = getSheet('謝金計算結果');
  const data = calcSheet.getDataRange().getValues();
  const headers = data[0];
  const year = params.year ? parseInt(params.year) : null;
  const month = params.month ? parseInt(params.month) : null;

  const rows = data.slice(1)
    .filter(row => {
      if (!year && !month) return true;
      const ym = String(row[2]); // 対象年月 "YYYY年MM月"
      if (year && !ym.includes(String(year))) return false;
      if (month && !ym.includes(String(month) + '月')) return false;
      return true;
    })
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    });

  return { success: true, data: rows };
}
