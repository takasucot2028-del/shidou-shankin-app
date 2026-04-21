/**
 * 謝金計算処理
 * CLAUDE.md の計算ロジックに従い、指導者・年月を指定して謝金を計算する
 */

// 時給設定
const HOURLY_RATE = {
  'メイン': 1600,
  'サブ': 1100,
};

// 源泉徴収率
const WITHHOLDING_TAX_RATE = 0.1021;

// 区分別上限時間（時間）
const MAX_HOURS = {
  '平日': 2,
  '休日': 3,
  '長期休暇': 3,
  '大会引率': 4,
};

// 教員の勤務終了時刻（分換算）16:40
const TEACHER_WORK_END_MINUTES = 16 * 60 + 40;

// ========== 謝金計算メイン ==========

/**
 * 謝金計算を実行してシートに保存する
 * body: { instructorName, year, month }
 * skipDelete=true の場合、シート上の既存行削除をスキップする（calcAllFees から呼ぶ場合）
 */
function calcFee(body, skipDelete) {
  const year = parseInt(body.year);
  const month = parseInt(body.month);
  const instructorName = body.instructorName;

  Logger.log('[calcFee] 開始: 指導者="%s", year=%s, month=%s, skipDelete=%s', instructorName, year, month, !!skipDelete);

  // 月報データを取得
  const reportSheet = getOrCreateReportSheet(year);
  const reportData = reportSheet.getDataRange().getValues().slice(1);

  // 対象指導者・年月・提出済みの行を絞り込む
  const rows = reportData.filter(row =>
    String(row[COL.INSTRUCTOR_NAME]).trim() === String(instructorName).trim() &&
    parseInt(row[COL.YEAR]) === year &&
    parseInt(row[COL.MONTH]) === month &&
    row[COL.STATUS] === '提出済'
  );

  Logger.log('[calcFee] 対象月報行数: %s', rows.length);

  if (rows.length === 0) {
    return { success: false, error: '対象の提出済み月報が見つかりません' };
  }

  // 指導者マスタから種別を取得
  const instructor = findInstructor(instructorName);
  if (!instructor) {
    return { success: false, error: '指導者マスタに登録がありません: ' + instructorName };
  }

  const result = computeFee(rows, instructor);
  saveFeeResult(instructorName, year, month, result, instructor.clubName, skipDelete);

  return { success: true, result };
}

/**
 * 謝金計算結果を手動修正する
 * body: { calcId, overrides: { fee, withholding, netPay } }
 */
function updateFee(body) {
  if (!body.calcId) throw new Error('calcId が必要です');

  const sheet = getSheet('謝金計算結果');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(body.calcId)) {
      const overrides = body.overrides || {};
      if (overrides.fee !== undefined) sheet.getRange(i + 1, 10).setValue(overrides.fee);
      if (overrides.withholding !== undefined) sheet.getRange(i + 1, 11).setValue(overrides.withholding);
      if (overrides.netPay !== undefined) sheet.getRange(i + 1, 12).setValue(overrides.netPay);
      // 修正フラグを立てる
      sheet.getRange(i + 1, 15).setValue(true);
      return { success: true };
    }
  }

  return { success: false, error: '計算IDが見つかりません: ' + body.calcId };
}

// ========== 計算ロジック ==========

/**
 * 1指導者分の月間謝金を計算する
 */
function computeFee(rows, instructor) {
  let totalCalcHours = 0;
  let mainCalcHours = 0;
  let subCalcHours = 0;
  let travelTotal = 0;

  // 区分別集計用
  const categoryHours = {
    '平日': 0,
    '休日': 0,
    '長期休暇': 0,
    '大会引率': 0,
  };

  rows.forEach(row => {
    const category = row[COL.CATEGORY];
    const rateType = row[COL.RATE_TYPE];
    const startTime = row[COL.START_TIME];
    const endTime = row[COL.END_TIME];
    const travelAmount = parseFloat(row[COL.TRAVEL_AMOUNT]) || 0;

    const calcHours = calcInstructionHours(
      startTime,
      endTime,
      category,
      instructor.type
    );

    // 行の謝金計算時間を更新
    const reportSheet = getOrCreateReportSheet(parseInt(row[COL.YEAR]));
    updateCalcHoursInSheet(reportSheet, row[COL.SUBMIT_ID], row[COL.DATE], calcHours);

    categoryHours[category] = (categoryHours[category] || 0) + calcHours;
    totalCalcHours += calcHours;

    if (rateType === 'メイン') {
      mainCalcHours += calcHours;
    } else {
      subCalcHours += calcHours;
    }

    travelTotal += travelAmount;
  });

  // 謝金額計算
  const fee = Math.round(mainCalcHours * HOURLY_RATE['メイン'] + subCalcHours * HOURLY_RATE['サブ']);
  const withholding = Math.floor(fee * WITHHOLDING_TAX_RATE);
  const netPay = fee - withholding;

  return {
    categoryHours,
    mainCalcHours,
    subCalcHours,
    totalCalcHours,
    fee,
    withholding,
    netPay,
    travelTotal,
  };
}

/**
 * 1行分の謝金計算時間を求める（15分単位切捨）
 */
function calcInstructionHours(startTime, endTime, category, instructorType) {
  const startMin = timeToMinutes(startTime);
  const endMin = timeToMinutes(endTime);

  if (startMin === null || endMin === null || endMin <= startMin) return 0;

  let effectiveMinutes;

  // 教員かつ平日・長期休暇は16:40以降のみ有効
  const isTeacher = instructorType === '教員';
  const isWeekdayOrLongVacation = category === '平日' || category === '長期休暇';

  if (isTeacher && isWeekdayOrLongVacation) {
    const effectiveStart = Math.max(startMin, TEACHER_WORK_END_MINUTES);
    effectiveMinutes = Math.max(endMin - effectiveStart, 0);
  } else {
    effectiveMinutes = endMin - startMin;
  }

  const maxMinutes = (MAX_HOURS[category] || 0) * 60;
  const cappedMinutes = Math.min(effectiveMinutes, maxMinutes);

  // 15分単位切捨
  const flooredMinutes = Math.floor(cappedMinutes / 15) * 15;

  return flooredMinutes / 60;
}

/**
 * 時刻値を分に変換する
 * GASはシートの時刻セルをDateオブジェクト・小数・"HH:MM"文字列のいずれかで返すため全形式に対応する
 */
function timeToMinutes(timeStr) {
  if (timeStr === null || timeStr === undefined || timeStr === '') return null;
  // Dateオブジェクト（GASがシートの時刻セルを変換した場合）
  if (timeStr instanceof Date) {
    return timeStr.getHours() * 60 + timeStr.getMinutes();
  }
  const str = String(timeStr).trim();
  // 小数フォーマット（例: 0.375 = 09:00、Sheetsの内部時刻表現）
  const num = parseFloat(str);
  if (!isNaN(num) && num >= 0 && num < 1) {
    return Math.round(num * 24 * 60);
  }
  // "HH:MM" 文字列
  const match = str.match(/^(\d{1,2}):(\d{2})$/);
  if (!match) return null;
  return parseInt(match[1]) * 60 + parseInt(match[2]);
}

// ========== シート保存・更新 ==========

/**
 * 謝金計算結果シートに保存（同一指導者・年月の既存行を削除して再挿入）
 * skipDelete=true の場合、削除をスキップして挿入のみ行う（calcAllFees の pre-delete 後に使用）
 */
function saveFeeResult(instructorName, year, month, result, clubName, skipDelete) {
  const sheet = getOrCreateFeeSheet();
  const ymLabel = year + '年' + month + '月';

  if (skipDelete) {
    Logger.log('[saveFeeResult] 削除スキップ（呼び出し元で一括削除済）: 指導者="%s", 年月="%s"', instructorName, ymLabel);
  } else {
    // 既存レコードを後ろから削除
    // Date型・ISO文字列・"YYYY年M月"文字列・数値シリアルすべてに対応
    const data = sheet.getDataRange().getValues();
    let deletedCount = 0;
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][1]).trim() === String(instructorName).trim() &&
          matchYearMonth(data[i][2], year, month)) {
        Logger.log('[saveFeeResult] 削除: row=%s, 氏名="%s", 年月セル値="%s"', i + 1, data[i][1], data[i][2]);
        sheet.deleteRow(i + 1);
        deletedCount++;
      }
    }
    Logger.log('[saveFeeResult] 削除完了: 指導者="%s", 年月="%s", 削除件数=%s', instructorName, ymLabel, deletedCount);
  }

  // 新規挿入（C列はテキスト形式を強制してSheetsの日付自動変換を防ぐ）
  const calcId = generateUUID();
  const lastRow = sheet.getLastRow() + 1;
  Logger.log('[saveFeeResult] 新規挿入: row=%s, 指導者="%s", 年月="%s"', lastRow, instructorName, ymLabel);
  writeFeeRow(sheet, lastRow, calcId, instructorName, ymLabel, result);
}

function writeFeeRow(sheet, rowIndex, calcId, instructorName, ymLabel, result) {
  const values = [[
    calcId,
    instructorName,
    ymLabel,
    result.categoryHours['平日'] || 0,
    result.categoryHours['休日'] || 0,
    result.categoryHours['長期休暇'] || 0,
    result.categoryHours['大会引率'] || 0,
    result.mainCalcHours,
    result.subCalcHours,
    result.fee,
    result.withholding,
    result.netPay,
    result.travelTotal,
    new Date(),
    false,
  ]];
  sheet.getRange(rowIndex, 1, 1, values[0].length).setValues(values);
  // C列（対象年月）をテキスト形式に強制してSheetsの日付自動変換を防ぐ
  sheet.getRange(rowIndex, 3).setNumberFormat('@STRING@');
}

function getOrCreateFeeSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName('謝金計算結果');
  if (!sheet) {
    sheet = ss.insertSheet('謝金計算結果');
    const headers = [
      '計算ID', '指導者氏名', '対象年月',
      '平日謝金計算時間', '休日謝金計算時間', '長期休暇謝金計算時間', '大会引率謝金計算時間',
      'メイン単価適用時間', 'サブ単価適用時間',
      '謝金総額', '源泉徴収額', '差引支払額', '旅費総額',
      '計算日時', '修正フラグ',
    ];
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    // C列（対象年月）をテキスト形式に設定
    sheet.getRange(1, 3, 1000, 1).setNumberFormat('@STRING@');
  }
  return sheet;
}

/**
 * 月報シートの謝金計算時間列を更新する
 */
function updateCalcHoursInSheet(sheet, submitId, date, calcHours) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][COL.SUBMIT_ID] === submitId && String(data[i][COL.DATE]) === String(date)) {
      sheet.getRange(i + 1, COL.CALC_HOURS + 1).setValue(calcHours);
      break;
    }
  }
}

// ========== 全指導者一括謝金計算 ==========

/**
 * 指定年月の提出済み月報がある全指導者の謝金を計算・保存して返す
 * body: { year, month }
 */
function calcAllFees(body) {
  const year = parseInt(body.year);
  const month = parseInt(body.month);

  const reportSheet = getOrCreateReportSheet(year);
  const reportData = reportSheet.getDataRange().getValues().slice(1);

  // 対象年月・提出済みの指導者名を重複なく収集
  const nameSet = {};
  reportData.forEach(row => {
    if (parseInt(row[COL.YEAR]) === year &&
        parseInt(row[COL.MONTH]) === month &&
        row[COL.STATUS] === '提出済') {
      nameSet[row[COL.INSTRUCTOR_NAME]] = true;
    }
  });

  const names = Object.keys(nameSet);
  if (names.length === 0) {
    return { success: true, data: [], message: '対象年月に提出済み月報がありません' };
  }

  // 計算前に対象年月の既存謝金計算結果を一括削除して重複を防ぐ
  // ここで全件削除した後、calcFee は skipDelete=true で呼び出して二重削除を防ぐ
  const feeSheet = getOrCreateFeeSheet();
  const feeData = feeSheet.getDataRange().getValues();
  Logger.log('[calcAllFees] 一括削除開始: year=%s, month=%s, 現在の総行数(ヘッダー含む)=%s', year, month, feeData.length);
  let bulkDeleted = 0;
  for (let i = feeData.length - 1; i >= 1; i--) {
    if (matchYearMonth(feeData[i][2], year, month)) {
      Logger.log('[calcAllFees] 削除: row=%s, 氏名="%s", 年月セル値="%s"', i + 1, feeData[i][1], feeData[i][2]);
      feeSheet.deleteRow(i + 1);
      bulkDeleted++;
    }
  }
  Logger.log('[calcAllFees] 一括削除完了: 削除件数=%s, 残行数(ヘッダー含む)=%s', bulkDeleted, feeSheet.getLastRow());

  const results = [];
  names.forEach(name => {
    try {
      // skipDelete=true: 上の一括削除で既存行は全て消えているため、saveFeeResult 内の削除をスキップ
      const res = calcFee({ instructorName: name, year, month }, true);
      if (res.success) {
        const instructor = findInstructor(name) || {};
        results.push({
          '指導者氏名': name,
          'クラブ名': instructor.clubName || '',
          '平日謝金計算時間': res.result.categoryHours['平日'] || 0,
          '休日謝金計算時間': res.result.categoryHours['休日'] || 0,
          '長期休暇謝金計算時間': res.result.categoryHours['長期休暇'] || 0,
          '大会引率謝金計算時間': res.result.categoryHours['大会引率'] || 0,
          '謝金総額': res.result.fee,
          '源泉徴収額': res.result.withholding,
          '差引支払額': res.result.netPay,
          '旅費総額': res.result.travelTotal,
          '修正フラグ': false,
        });
      }
    } catch (e) {
      // 個別エラーはスキップして続行
    }
  });

  return { success: true, data: results };
}

// findInstructor / getInstructors / updateMaster は Master.gs に定義

// ========== 口座振替データ生成 ==========

/**
 * 指定年月の謝金計算結果と指導者マスタを結合して「口座振替データ」シートを生成する
 * body: { year, month }
 */
function generateTransferSheet(body) {
  const year = parseInt(body.year);
  const month = parseInt(body.month);
  const ymLabel = year + '年' + month + '月';

  Logger.log('[generateTransferSheet] 受信パラメータ: body.year=%s, body.month=%s', body.year, body.month);
  Logger.log('[generateTransferSheet] parseInt後: year=%s, month=%s', year, month);
  Logger.log('[generateTransferSheet] フィルタ用 ymLabel="%s"', ymLabel);

  // 謝金計算結果から対象年月の行を取得（なければ自動計算して再取得）
  const feeSheet = getOrCreateFeeSheet();
  const allValues = feeSheet.getDataRange().getValues();
  Logger.log('[generateTransferSheet] 謝金計算結果シート 総行数(ヘッダー含む)=%s', allValues.length);
  if (allValues.length > 0) {
    Logger.log('[generateTransferSheet] 1行目(ヘッダー): %s', JSON.stringify(allValues[0]));
  }
  if (allValues.length > 1) {
    Logger.log('[generateTransferSheet] 2行目(データ): %s', JSON.stringify(allValues[1]));
    Logger.log('[generateTransferSheet] 2行目C列(対象年月)の値="%s", type=%s', allValues[1][2], typeof allValues[1][2]);
  }

  let feeData = allValues.slice(1);
  let feeRows = feeData.filter(row => matchYearMonth(row[2], year, month));
  Logger.log('[generateTransferSheet] ymLabel="%s" でフィルタ後の件数=%s', ymLabel, feeRows.length);

  if (feeRows.length === 0) {
    const calcResult = calcAllFees({ year: year, month: month });
    if (!calcResult.success || calcResult.data.length === 0) {
      return { success: false, error: '対象年月の提出済み月報がありません: ' + ymLabel };
    }
    feeData = feeSheet.getDataRange().getValues().slice(1);
    feeRows = feeData.filter(row => matchYearMonth(row[2], year, month));
  }

  if (feeRows.length === 0) {
    return { success: false, error: '対象年月の謝金計算結果がありません: ' + ymLabel };
  }

  // 指導者マスタをマップ化（氏名 → マスタ行）
  const masterSheet = getSheet('指導者マスタ');
  const masterData = masterSheet.getDataRange().getValues().slice(1);
  const masterMap = {};
  masterData.forEach(row => {
    if (row[1]) masterMap[row[1]] = row;
  });

  // 差引支払額 > 0 の行のみ抽出・結合
  const outputRows = [];
  feeRows.forEach(row => {
    const netPay = parseFloat(row[11]) || 0;
    if (netPay <= 0) return;

    const name = row[1];
    const master = masterMap[name];
    outputRows.push([
      '',                          // A: No（後で連番付与）
      name,                        // B: 指導者氏名
      master ? master[2] : '',     // C: クラブ名
      master ? master[3] : '',     // D: 区分
      ymLabel,                     // E: 対象年月
      parseFloat(row[9]) || 0,     // F: 謝金総額
      parseFloat(row[10]) || 0,    // G: 源泉徴収額
      netPay,                      // H: 差引支払額
      master ? master[6] : '',     // I: 金融機関名
      master ? master[7] : '',     // J: 支店名
      master ? master[8] : '',     // K: 口座種別
      master ? master[9] : '',     // L: 口座番号
    ]);
  });

  if (outputRows.length === 0) {
    return { success: false, error: '差引支払額が0円より大きい指導者がいません' };
  }

  // 連番付与
  outputRows.forEach((row, i) => { row[0] = i + 1; });

  // シートを取得または作成して上書き
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName('口座振替データ');
  if (!sheet) {
    sheet = ss.insertSheet('口座振替データ');
  } else {
    sheet.clearContents();
  }

  const headers = [
    'No', '指導者氏名', 'クラブ名', '区分', '対象年月',
    '謝金総額（円）', '源泉徴収額（円）', '差引支払額（円）',
    '金融機関名', '支店名', '口座種別', '口座番号',
  ];
  sheet.appendRow(headers);
  sheet.setFrozenRows(1);

  if (outputRows.length > 0) {
    sheet.getRange(2, 1, outputRows.length, headers.length).setValues(outputRows);
  }

  return { success: true, count: outputRows.length, sheetName: '口座振替データ' };
}

// ========== 年月比較ユーティリティ ==========

/**
 * 謝金計算結果シートのC列値（対象年月）と指定の年月を比較する。
 * 以下の全形式に対応:
 *   - "2026年4月" 形式の文字列（テキスト保存時の正規形式）
 *   - Dateオブジェクト（GASがセル値を日付型に変換した場合）
 *   - "2026-04-01T..." ISO文字列
 *   - 数値（Sheetsの日付シリアル値。GASでは通常Dateになるが念のため対応）
 */
function matchYearMonth(cellValue, year, month) {
  if (cellValue === null || cellValue === undefined || cellValue === '') return false;

  // Dateオブジェクト
  if (cellValue instanceof Date) {
    if (isNaN(cellValue.getTime())) return false;
    return cellValue.getFullYear() === year && (cellValue.getMonth() + 1) === month;
  }

  // 数値（Sheetsの日付シリアル値）→ Dateに変換して比較
  // Sheetsシリアル値: 1900-01-01 = 1, 以降1日=1
  if (typeof cellValue === 'number') {
    const msPerDay = 86400000;
    const epoch = new Date(Date.UTC(1899, 11, 30)).getTime();
    const d = new Date(epoch + cellValue * msPerDay);
    return d.getUTCFullYear() === year && (d.getUTCMonth() + 1) === month;
  }

  const str = String(cellValue).trim();

  // "YYYY年M月" または "YYYY年MM月" 形式
  const jpMatch = str.match(/^(\d{4})年(\d{1,2})月$/);
  if (jpMatch) {
    return parseInt(jpMatch[1]) === year && parseInt(jpMatch[2]) === month;
  }

  // ISO文字列 "YYYY-MM-..." 形式
  const isoMatch = str.match(/^(\d{4})-(\d{2})/);
  if (isoMatch) {
    return parseInt(isoMatch[1]) === year && parseInt(isoMatch[2]) === month;
  }

  return false;
}

/**
 * 謝金計算結果シートのC列をDate型から文字列形式（"YYYY年M月"）に一括変換する。
 * 既存データ修正用。Apps Scriptエディタから手動実行すること。
 */
function fixFeeSheetYearMonth() {
  const sheet = getOrCreateFeeSheet();
  const data = sheet.getDataRange().getValues();
  let fixed = 0;
  for (let i = 1; i < data.length; i++) {
    const cell = data[i][2];
    if (cell instanceof Date) {
      const label = cell.getFullYear() + '年' + (cell.getMonth() + 1) + '月';
      sheet.getRange(i + 1, 3).setValue(label);
      fixed++;
    }
  }
  Logger.log('fixFeeSheetYearMonth: %s 行を修正しました', fixed);
}

// ========== 重複データクリーンアップ ==========

/**
 * 謝金計算結果シートの重複行を削除する。
 * 同一の指導者氏名・対象年月の組み合わせは計算日時が最新の1件だけ残す。
 * 戻り値: { deleted: 削除件数, kept: 残存件数 }
 */
function cleanupFeeResults() {
  const sheet = getOrCreateFeeSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { deleted: 0, kept: 0 };

  // キー(氏名+年月) → 最新行インデックス(1-based)を収集
  const latestRow = {}; // key → { rowIndex, calcDate }
  for (let i = 1; i < data.length; i++) {
    const name = data[i][1];
    const ymRaw = data[i][2];
    const calcDate = data[i][13] instanceof Date ? data[i][13] : new Date(data[i][13] || 0);
    const key = name + '|' + (ymRaw instanceof Date
      ? ymRaw.getFullYear() + '年' + (ymRaw.getMonth() + 1) + '月'
      : String(ymRaw));
    if (!latestRow[key] || calcDate > latestRow[key].calcDate) {
      latestRow[key] = { rowIndex: i + 1, calcDate };
    }
  }

  const keepRows = new Set(Object.values(latestRow).map(v => v.rowIndex));

  // 後ろから削除してインデックスずれを防ぐ
  let deleted = 0;
  for (let i = data.length; i >= 2; i--) {
    if (!keepRows.has(i)) {
      sheet.deleteRow(i);
      deleted++;
    }
  }

  return { deleted, kept: keepRows.size };
}

/**
 * Apps Scriptエディタから手動実行して重複データを一括削除する。
 * 実行後はログで削除件数を確認すること。
 */
function runCleanupFeeResults() {
  Logger.log('=== runCleanupFeeResults 開始 ===');
  const result = cleanupFeeResults();
  Logger.log('削除件数: %s, 残存件数: %s', result.deleted, result.kept);
  Logger.log('=== runCleanupFeeResults 完了 ===');
}

// ========== デバッグ用テスト関数 ==========

/**
 * 教員・平日・16:40開始・18:10終了のケースで謝金計算時間を検証する
 * Apps Scriptエディタから手動実行し、ログで結果を確認すること
 */
function testCalcFee1640() {
  Logger.log('=== testCalcFee1640 開始 ===');

  const startTime = '16:40';
  const endTime   = '18:10';
  const category  = '平日';
  const instructorType = '教員';

  // 入力値の分換算を確認
  const startMin = timeToMinutes(startTime);
  const endMin   = timeToMinutes(endTime);
  Logger.log('開始時刻: %s → %s 分', startTime, startMin);
  Logger.log('終了時刻: %s → %s 分', endTime,   endMin);
  Logger.log('TEACHER_WORK_END_MINUTES: %s 分 (%s)', TEACHER_WORK_END_MINUTES,
             Math.floor(TEACHER_WORK_END_MINUTES / 60) + ':' + ('0' + (TEACHER_WORK_END_MINUTES % 60)).slice(-2));

  // 有効開始時刻
  const effectiveStart   = Math.max(startMin, TEACHER_WORK_END_MINUTES);
  const overlapMinutes   = effectiveStart - startMin;   // 勤務重複時間（分）
  const effectiveMinutes = Math.max(endMin - effectiveStart, 0);
  const maxMinutes       = (MAX_HOURS[category] || 0) * 60;
  const cappedMinutes    = Math.min(effectiveMinutes, maxMinutes);
  const flooredMinutes   = Math.floor(cappedMinutes / 15) * 15;

  Logger.log('--- 計算過程 ---');
  Logger.log('有効開始時刻 (effectiveStart): %s 分', effectiveStart);
  Logger.log('勤務重複時間 (overlapMinutes): %s 分 ← 期待値: 0分', overlapMinutes);
  Logger.log('有効指導時間 (effectiveMinutes): %s 分 ← 期待値: 90分 (1h30m)', effectiveMinutes);
  Logger.log('上限 (maxMinutes): %s 分 (%s h)', maxMinutes, MAX_HOURS[category]);
  Logger.log('上限適用後 (cappedMinutes): %s 分', cappedMinutes);
  Logger.log('15分切捨後 (flooredMinutes): %s 分 ← 期待値: 90分 (1h30m)', flooredMinutes);

  // calcInstructionHours 経由でも確認
  const calcHours = calcInstructionHours(startTime, endTime, category, instructorType);
  Logger.log('--- calcInstructionHours の戻り値 ---');
  Logger.log('calcHours: %s 時間 ← 期待値: 1.5時間', calcHours);

  // 判定
  const PASS = '✓ PASS';
  const FAIL = '✗ FAIL';
  Logger.log('--- 判定 ---');
  Logger.log('勤務重複時間: %s', overlapMinutes === 0   ? PASS : FAIL + ' (実際: ' + overlapMinutes + '分)');
  Logger.log('実質指導時間: %s', effectiveMinutes === 90 ? PASS : FAIL + ' (実際: ' + effectiveMinutes + '分)');
  Logger.log('15分切捨後:   %s', flooredMinutes === 90  ? PASS : FAIL + ' (実際: ' + flooredMinutes + '分)');
  Logger.log('calcHours:    %s', calcHours === 1.5       ? PASS : FAIL + ' (実際: ' + calcHours + '時間)');

  Logger.log('=== testCalcFee1640 完了 ===');
}

/**
 * Apps Scriptエディタから手動実行して口座振替シート生成をテストする
 * ★ 実行前に year / month を実際のデータに合わせて変更すること
 */
function testGenerateTransfer() {
  const year  = 2026; // ← テストしたい年に変更
  const month = 4;    // ← テストしたい月に変更

  Logger.log('=== testGenerateTransfer 開始: year=%s, month=%s ===', year, month);

  // 1. 謝金計算結果シートの全データを確認
  const feeSheet = getOrCreateFeeSheet();
  const allValues = feeSheet.getDataRange().getValues();
  Logger.log('謝金計算結果シート 総行数(ヘッダー含む): %s', allValues.length);
  if (allValues.length > 0) Logger.log('ヘッダー行: %s', JSON.stringify(allValues[0]));
  for (let i = 1; i < Math.min(allValues.length, 6); i++) {
    Logger.log('データ行%s: C列(対象年月)="%s", 全体=%s', i, allValues[i][2], JSON.stringify(allValues[i]));
  }

  // 2. フィルタ確認
  const ymLabel = year + '年' + month + '月';
  Logger.log('検索する ymLabel="%s"', ymLabel);
  const matched = allValues.slice(1).filter(row => matchYearMonth(row[2], year, month));
  Logger.log('ymLabel一致件数: %s', matched.length);

  // 3. generateTransferSheet を実行
  const result = generateTransferSheet({ year: year, month: month });
  Logger.log('=== generateTransferSheet 結果: %s ===', JSON.stringify(result));
}
