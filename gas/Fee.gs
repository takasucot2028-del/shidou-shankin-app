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

// 教員の勤務終了時刻（分換算）16:45
const TEACHER_WORK_END_MINUTES = 16 * 60 + 45;

// ========== 謝金計算メイン ==========

/**
 * 謝金計算を実行してシートに保存する
 * body: { instructorName, year, month }
 */
function calcFee(body) {
  const year = parseInt(body.year);
  const month = parseInt(body.month);
  const instructorName = body.instructorName;

  // 月報データを取得
  const reportSheet = getOrCreateReportSheet(year);
  const reportData = reportSheet.getDataRange().getValues().slice(1);

  // 対象指導者・年月・提出済みの行を絞り込む
  const rows = reportData.filter(row =>
    row[COL.INSTRUCTOR_NAME] === instructorName &&
    parseInt(row[COL.YEAR]) === year &&
    parseInt(row[COL.MONTH]) === month &&
    row[COL.STATUS] === '提出済'
  );

  if (rows.length === 0) {
    return { success: false, error: '対象の提出済み月報が見つかりません' };
  }

  // 指導者マスタから種別を取得
  const instructor = findInstructor(instructorName);
  if (!instructor) {
    return { success: false, error: '指導者マスタに登録がありません: ' + instructorName };
  }

  const result = computeFee(rows, instructor);
  saveFeeResult(instructorName, year, month, result, instructor.clubName);

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

  // 教員かつ平日・長期休暇は16:45以降のみ有効
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
 */
function saveFeeResult(instructorName, year, month, result, clubName) {
  const sheet = getOrCreateFeeSheet();
  const ymLabel = year + '年' + month + '月';

  // 既存レコードを後ろから削除してインデックスずれを防ぐ
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][1] === instructorName && data[i][2] === ymLabel) {
      sheet.deleteRow(i + 1);
    }
  }

  // 新規追加
  const calcId = generateUUID();
  const lastRow = sheet.getLastRow() + 1;
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
    false, // 修正フラグ
  ]];
  sheet.getRange(rowIndex, 1, 1, values[0].length).setValues(values);
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

  const results = [];
  names.forEach(name => {
    try {
      const res = calcFee({ instructorName: name, year, month });
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
  let feeRows = feeData.filter(row => row[2] === ymLabel);
  Logger.log('[generateTransferSheet] ymLabel="%s" でフィルタ後の件数=%s', ymLabel, feeRows.length);

  if (feeRows.length === 0) {
    const calcResult = calcAllFees({ year: year, month: month });
    if (!calcResult.success || calcResult.data.length === 0) {
      return { success: false, error: '対象年月の提出済み月報がありません: ' + ymLabel };
    }
    feeData = feeSheet.getDataRange().getValues().slice(1);
    feeRows = feeData.filter(row => row[2] === ymLabel);
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

// ========== デバッグ用テスト関数 ==========

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
  const matched = allValues.slice(1).filter(row => row[2] === ymLabel);
  Logger.log('ymLabel一致件数: %s', matched.length);

  // 3. generateTransferSheet を実行
  const result = generateTransferSheet({ year: year, month: month });
  Logger.log('=== generateTransferSheet 結果: %s ===', JSON.stringify(result));
}
