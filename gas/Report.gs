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
    if (params.instructorName && String(row[COL.INSTRUCTOR_NAME]).trim() !== String(params.instructorName).trim()) return false;
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
    date: sheetDateToString(row[COL.DATE]),
    category: row[COL.CATEGORY],
    rateType: row[COL.RATE_TYPE],
    startTime: sheetTimeToString(row[COL.START_TIME]),
    endTime: sheetTimeToString(row[COL.END_TIME]),
    instructionHours: row[COL.INSTRUCTION_HOURS],
    calcHours: row[COL.CALC_HOURS],
    transport: row[COL.TRANSPORT],
    destination: row[COL.DESTINATION],
    travelAmount: row[COL.TRAVEL_AMOUNT],
    note: row[COL.NOTE],
    status: String(row[COL.STATUS] || '').trim(),
    submittedAt: row[COL.SUBMITTED_AT],
    updatedAt: row[COL.UPDATED_AT],
  };
}

function sheetDateToString(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(val);
}

function sheetTimeToString(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'HH:mm');
  }
  return String(val);
}

// ========== 排他ロック ==========

/**
 * 月報の書き込み処理を排他制御する。
 * 複数の保存/提出が同時に実行されると「削除→挿入」の間に競合し、
 * 行が二重挿入されるため、スクリプトロックで直列化する。
 */
function withReportLock(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000); // 最大20秒待機
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

// ========== 月報一時保存 ==========

/**
 * 月報一時保存（下書き）
 * body: { instructorName, clubName, year, month, rows: [...] }
 * 同じ指導者・年月の既存下書き行を削除して再挿入する
 */
function saveReport(body) {
  validateReportBody(body);

  return withReportLock(() => {
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

    SpreadsheetApp.flush();
    return { success: true, submitId };
  });
}

// ========== 月報提出 ==========

/**
 * 月報提出（ステータスを「提出済」に変更）
 * body: { instructorName, clubName, year, month, rows: [...], submitId }
 * 既存の下書きを削除して提出済み行を挿入する
 */
function submitReport(body) {
  validateReportBody(body);

  const result = withReportLock(() => {
    const year = parseInt(body.year);
    const month = parseInt(body.month);
    const sheet = getOrCreateReportSheet(year);

    // 既存の提出済み行を削除（再提出を許可）
    deleteReportRows(sheet, body.instructorName, year, month, '提出済');
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

    // setValuesのバッファをシートにコミットしてから謝金計算を実行。
    // forceRecalc=true：事務局の手修正フラグが立っていても再計算する。
    // ※再提出で内容が変わっているのに旧金額が残ると、払い漏れ・過大払いの
    //   どちらも無警告で起きるため、月報の内容を常に正とする。
    SpreadsheetApp.flush();
    const feeRes = calcFee({
      instructorName: body.instructorName, year, month, forceRecalc: true,
    });
    if (!feeRes || !feeRes.success) {
      Logger.log('[submitReport] 警告: 謝金計算に失敗しました: %s',
        feeRes ? feeRes.error : '(戻り値なし)');
    }

    return { submitId, now, feeError: (feeRes && feeRes.success) ? null : (feeRes ? feeRes.error : '謝金計算に失敗しました') };
  });

  // 事務局へ提出通知メール（ロック外で実行）
  notifyAdminOnSubmit(body.instructorName, body.clubName, parseInt(body.year), parseInt(body.month), result.now);

  if (result.feeError) {
    return { success: true, submitId: result.submitId,
      warning: '月報は提出しましたが謝金の計算に失敗しました: ' + result.feeError };
  }
  return { success: true, submitId: result.submitId };
}

// ========== 月報修正（事務局） ==========

/**
 * 事務局による月報修正（提出済みデータを上書き・謝金再計算）
 * body: { instructorName, clubName, year, month, submitId, rows: [...] }
 */
function updateReport(body) {
  validateReportBody(body);

  return withReportLock(() => {
    const year = parseInt(body.year);
    const month = parseInt(body.month);
    const sheet = getOrCreateReportSheet(year);

    deleteReportRows(sheet, body.instructorName, year, month, '提出済');
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
      submittedAt: body.submittedAt || now,
      updatedAt: now,
    }));

    if (newRows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length)
        .setValues(newRows);
    }

    // forceRecalc=true：手修正フラグを無視して再計算する（Dの方針）
    SpreadsheetApp.flush();
    const feeRes = calcFee({
      instructorName: body.instructorName, year, month, forceRecalc: true,
    });
    if (!feeRes || !feeRes.success) {
      Logger.log('[updateReport] 警告: 謝金計算に失敗しました: %s',
        feeRes ? feeRes.error : '(戻り値なし)');
      return { success: true, submitId, warning: '月報は保存しましたが謝金の再計算に失敗しました: ' +
        (feeRes ? feeRes.error : '不明なエラー') };
    }

    return { success: true, submitId };
  });
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
    name: String(r[1]).trim(),
    clubName: r[2],
    email: r[10],
  }));

  // 対象月の提出済み指導者を集計
  const submittedMap = {};
  data.forEach(row => {
    if (parseInt(row[COL.YEAR]) === year &&
        parseInt(row[COL.MONTH]) === month &&
        String(row[COL.STATUS]).trim() === '提出済') {
      submittedMap[String(row[COL.INSTRUCTOR_NAME]).trim()] = {
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
  // 氏名・ステータスはtrimして比較する。
  // ※厳密一致にすると、セルに空白が混入した行が削除されずに残り、
  //   再提出時に旧行＋新行が併存して謝金が二重計上される。
  const targetName   = String(instructorName).trim();
  const targetStatus = String(status).trim();
  // 後ろから削除してインデックスずれを防ぐ
  for (let i = data.length - 1; i >= 1; i--) {
    const row = data[i];
    if (String(row[COL.INSTRUCTOR_NAME]).trim() === targetName &&
        parseInt(row[COL.YEAR]) === year &&
        parseInt(row[COL.MONTH]) === month &&
        String(row[COL.STATUS]).trim() === targetStatus) {
      sheet.deleteRow(i + 1); // シートは1-indexed
    }
  }
}

function hasSubmittedReport(sheet, instructorName, year, month) {
  const data = sheet.getDataRange().getValues().slice(1);
  return data.some(row =>
    String(row[COL.INSTRUCTOR_NAME]).trim() === String(instructorName).trim() &&
    parseInt(row[COL.YEAR]) === year &&
    parseInt(row[COL.MONTH]) === month &&
    String(row[COL.STATUS]).trim() === '提出済'
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
      // 年月両方ある場合は完全一致で判定（"1月"が"11月"に誤マッチするのを防ぐ）
      if (year && month) return ym === String(year) + '年' + String(month) + '月';
      if (year)  return ym.startsWith(String(year) + '年');
      if (month) return ym === String(month) + '月' || ym.endsWith('年' + String(month) + '月');
      return true;
    })
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    });

  return { success: true, data: rows };
}

// ========== 重複行の検査・除去（保守用） ==========

/**
 * 月報データの重複判定キー。
 * 提出IDは含めない（別IDでも内容が同じなら重複とみなす）。
 *
 * ※同一指導者・同一日でも、メイン/サブの別行や時刻違いの別コマは
 *   正当な別行なので、区分・時給区分・開始・終了をキーに含める。
 */
function reportRowDuplicateKey(row) {
  return [
    String(row[COL.INSTRUCTOR_NAME]).trim(),
    parseInt(row[COL.YEAR]),
    parseInt(row[COL.MONTH]),
    sheetDateToString(row[COL.DATE]),
    String(row[COL.CATEGORY]).trim(),
    String(row[COL.RATE_TYPE]).trim(),
    sheetTimeToString(row[COL.START_TIME]),
    sheetTimeToString(row[COL.END_TIME]),
    String(row[COL.TRANSPORT]).trim(),
    String(row[COL.DESTINATION]).trim(),
    Number(row[COL.TRAVEL_AMOUNT]) || 0,
    String(row[COL.NOTE]).trim(),
    String(row[COL.STATUS]).trim(),
  ].join('|');
}

/**
 * 月報データの重複行を検査する（読み取り専用・シートは一切変更しない）
 * 実行後「実行ログ」で結果を確認する。
 */
function auditDuplicateReportRows(year) {
  const targetYear = parseInt(year) || new Date().getFullYear();
  const sheet = getOrCreateReportSheet(targetYear);
  const data = sheet.getDataRange().getValues();

  const seen = {};
  const groups = {};   // "氏名|年|月|ステータス" => 重複件数
  let duplicateTotal = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[COL.INSTRUCTOR_NAME]) continue;

    const key = reportRowDuplicateKey(row);
    if (seen[key]) {
      duplicateTotal++;
      const g = [
        String(row[COL.INSTRUCTOR_NAME]).trim(),
        parseInt(row[COL.YEAR]) + '年' + parseInt(row[COL.MONTH]) + '月',
        String(row[COL.STATUS]).trim(),
      ].join(' / ');
      groups[g] = (groups[g] || 0) + 1;
      Logger.log('重複: シート行%s  %s  %s', i + 1, g, sheetDateToString(row[COL.DATE]));
    } else {
      seen[key] = true;
    }
  }

  Logger.log('========== 検査結果 (%s) ==========', getReportSheetName(targetYear));
  Logger.log('総行数(ヘッダー除く): %s', Math.max(data.length - 1, 0));
  Logger.log('重複行の合計: %s 件', duplicateTotal);

  const summary = Object.keys(groups).sort().map(g => ({ group: g, count: groups[g] }));
  summary.forEach(s => Logger.log('  %s … 重複 %s 件', s.group, s.count));

  if (duplicateTotal === 0) Logger.log('重複はありません。');
  Logger.log('※この関数はシートを変更しません。削除は cleanupDuplicateReportRows を使用してください。');

  return { success: true, year: targetYear, duplicateTotal, summary };
}

/**
 * 月報データの重複行を削除する。
 *
 * 安全のため既定はドライラン（削除しない）。
 *   cleanupDuplicateReportRows(2026)        → 検査のみ（削除しない）
 *   cleanupDuplicateReportRows(2026, false) → 実際に削除する
 *
 * 各重複グループの「最初の1行」を残し、2件目以降を削除する。
 * 削除した場合は、影響のあった指導者・年月の謝金を強制再計算する
 * （修正フラグが立っていても、二重計上された金額を残さないため）。
 */
function cleanupDuplicateReportRows(year, dryRun) {
  const isDryRun = (dryRun !== false);  // 明示的に false を渡した場合のみ削除
  const targetYear = parseInt(year) || new Date().getFullYear();

  return withReportLock(() => {
    const sheet = getOrCreateReportSheet(targetYear);
    const data = sheet.getDataRange().getValues();

    const seen = {};
    const toDelete = [];      // 削除対象のシート行番号（1-indexed）
    const affected = {};      // "氏名|年|月" => true

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[COL.INSTRUCTOR_NAME]) continue;

      const key = reportRowDuplicateKey(row);
      if (seen[key]) {
        toDelete.push(i + 1);
        affected[[
          String(row[COL.INSTRUCTOR_NAME]).trim(),
          parseInt(row[COL.YEAR]),
          parseInt(row[COL.MONTH]),
        ].join('|')] = true;
        Logger.log('%s シート行%s: %s / %s / %s',
          isDryRun ? '[ドライラン] 削除対象' : '削除',
          i + 1,
          String(row[COL.INSTRUCTOR_NAME]).trim(),
          sheetDateToString(row[COL.DATE]),
          String(row[COL.STATUS]).trim());
      } else {
        seen[key] = true;
      }
    }

    const affectedList = Object.keys(affected).map(k => {
      const parts = k.split('|');
      return { instructorName: parts[0], year: parseInt(parts[1]), month: parseInt(parts[2]) };
    });

    Logger.log('========== %s ==========', isDryRun ? 'ドライラン結果' : '削除結果');
    Logger.log('対象シート: %s', getReportSheetName(targetYear));
    Logger.log('重複行: %s 件', toDelete.length);
    Logger.log('影響のある指導者・年月: %s 件', affectedList.length);
    affectedList.forEach(a => Logger.log('  %s %s年%s月', a.instructorName, a.year, a.month));

    if (isDryRun) {
      Logger.log('※ドライランのため削除していません。実行するには');
      Logger.log('   cleanupDuplicateReportRows(%s, false)', targetYear);
      return { success: true, dryRun: true, duplicateCount: toDelete.length, affected: affectedList };
    }

    // 後ろから削除してインデックスずれを防ぐ
    toDelete.sort((a, b) => b - a).forEach(rowNum => sheet.deleteRow(rowNum));
    SpreadsheetApp.flush();
    Logger.log('%s 行を削除しました。', toDelete.length);

    // 影響範囲の謝金を強制再計算（修正フラグを無視）
    affectedList.forEach(a => {
      const res = calcFee({
        instructorName: a.instructorName,
        year: a.year,
        month: a.month,
        forceRecalc: true,
      });
      Logger.log('再計算: %s %s年%s月 → %s',
        a.instructorName, a.year, a.month,
        res.success ? ('謝金=' + res.result.fee + '円 / 差引=' + res.result.netPay + '円') : ('失敗: ' + res.error));
    });

    return { success: true, dryRun: false, deletedCount: toDelete.length, affected: affectedList };
  });
}
