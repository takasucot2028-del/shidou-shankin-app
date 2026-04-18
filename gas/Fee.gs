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
 * "HH:MM" を分に変換する
 */
function timeToMinutes(timeStr) {
  if (!timeStr) return null;
  const str = String(timeStr);
  const match = str.match(/^(\d{1,2}):(\d{2})$/);
  if (!match) return null;
  return parseInt(match[1]) * 60 + parseInt(match[2]);
}

// ========== シート保存・更新 ==========

/**
 * 謝金計算結果シートに保存（既存レコードがあれば更新）
 */
function saveFeeResult(instructorName, year, month, result, clubName) {
  const sheet = getOrCreateFeeSheet();
  const data = sheet.getDataRange().getValues();
  const ymLabel = year + '年' + month + '月';

  // 既存レコード検索（修正フラグが立っていなければ上書き）
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === instructorName && data[i][2] === ymLabel && !data[i][14]) {
      writeFeeRow(sheet, i + 1, data[i][0], instructorName, ymLabel, result);
      return;
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

// findInstructor / getInstructors / updateMaster は Master.gs に定義
