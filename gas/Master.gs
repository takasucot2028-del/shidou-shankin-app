/**
 * 指導者マスタ処理
 * シート「指導者マスタ」の CRUD を担う
 *
 * 列定義（0-indexed）
 * 0:No 1:氏名 2:クラブ名 3:区分 4:指導者種別 5:時給
 * 6:金融機関 7:支店名 8:口座種別 9:口座番号 10:メールアドレス 11:登録日 12:PIN
 */

// ========== 取得 ==========

/**
 * 全指導者を返す（PINは除外）
 * GET ?action=getInstructors
 */
function getInstructors() {
  const sheet = getSheet('指導者マスタ');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1)
    .filter(r => r[1]) // 氏名が空の行を除外
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => {
        if (h !== 'PIN') obj[h] = row[i]; // PINはセキュリティのため除外
      });
      return obj;
    });
  return { success: true, data: rows };
}

/**
 * 氏名で指導者を1件返す（Fee.gs から参照）
 */
function findInstructor(name) {
  const sheet = getSheet('指導者マスタ');
  const data = sheet.getDataRange().getValues().slice(1);
  const row = data.find(r => r[1] === name);
  if (!row) return null;
  return {
    no: row[0],
    name: row[1],
    clubName: row[2],
    rateType: row[3],       // メイン / サブ
    type: row[4],           // 教員 / 一般
    hourlyRate: row[5],
    bank: row[6],
    branch: row[7],
    accountType: row[8],
    accountNumber: row[9],
    email: row[10],
  };
}

// ========== CRUD ==========

/**
 * 指導者マスタの追加・更新・削除
 * POST body: { action: 'updateMaster', operation: 'add'|'update'|'delete', ...fields }
 */
function updateMaster(body) {
  const sheet = getSheet('指導者マスタ');

  if (body.operation === 'add') {
    return addInstructor(sheet, body);
  }
  if (body.operation === 'update') {
    return editInstructor(sheet, body);
  }
  if (body.operation === 'delete') {
    return deleteInstructor(sheet, body);
  }

  return { success: false, error: '不正な operation: ' + body.operation };
}

function addInstructor(sheet, body) {
  validateMasterFields(body);
  if (body.pin && !/^\d{4}$/.test(String(body.pin))) {
    throw new Error('PINは4桁の数字で入力してください');
  }
  const data = sheet.getDataRange().getValues();
  const no = data.length; // ヘッダー行含む → 実質連番
  sheet.appendRow([
    no,
    body.name,
    body.clubName,
    body.rateType,
    body.instructorType,
    body.hourlyRate,
    body.bank        || '',
    body.branch      || '',
    body.accountType || '',
    body.accountNumber || '',
    body.email       || '',
    new Date(),
    body.pin         || '',
  ]);
  return { success: true };
}

function editInstructor(sheet, body) {
  validateMasterFields(body);
  if (body.pin && !/^\d{4}$/.test(String(body.pin))) {
    throw new Error('PINは4桁の数字で入力してください');
  }
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === body.name) {
      // PINが指定されていない場合は既存PINを保持
      const newPin = (body.pin !== undefined && body.pin !== '') ? String(body.pin) : String(data[i][12] || '');
      sheet.getRange(i + 1, 1, 1, 13).setValues([[
        data[i][0],
        body.name,
        body.clubName,
        body.rateType,
        body.instructorType,
        body.hourlyRate,
        body.bank        || '',
        body.branch      || '',
        body.accountType || '',
        body.accountNumber || '',
        body.email       || '',
        data[i][11], // 登録日は変えない
        newPin,
      ]]);
      return { success: true };
    }
  }
  return { success: false, error: '指導者が見つかりません: ' + body.name };
}

function deleteInstructor(sheet, body) {
  if (!body.name) return { success: false, error: '氏名が必要です' };
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][1] === body.name) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, error: '指導者が見つかりません: ' + body.name };
}

// ========== PIN照合 ==========

/**
 * 氏名とPINを照合する
 * POST body: { action: 'checkPin', name, pin }
 */
function checkPin(body) {
  if (!body.name || !body.pin) {
    return { success: false, error: '氏名とPINを入力してください' };
  }
  const sheet = getSheet('指導者マスタ');
  const data = sheet.getDataRange().getValues().slice(1);
  const row = data.find(r => r[1] === body.name);
  if (!row) {
    return { success: false, error: 'PINが正しくありません' };
  }
  const storedPin = String(row[12] || '');
  if (!storedPin) {
    return { success: false, error: 'PINが設定されていません。事務局にお問い合わせください' };
  }
  if (storedPin !== String(body.pin)) {
    return { success: false, error: 'PINが正しくありません' };
  }
  return { success: true, name: row[1] };
}

// ========== バリデーション ==========

function validateMasterFields(body) {
  if (!body.name)           throw new Error('氏名が必要です');
  if (!body.clubName)       throw new Error('クラブ名が必要です');
  if (!body.rateType)       throw new Error('区分（メイン/サブ）が必要です');
  if (!body.instructorType) throw new Error('指導者種別（教員/一般）が必要です');
  if (!body.hourlyRate)     throw new Error('時給が必要です');
}

// ========== マスタシート初期化 ==========

/**
 * 「指導者マスタ」シートがなければ作成してヘッダーを設定する
 * 初回セットアップ時に手動実行
 */
function initMasterSheet() {
  const ss = getSpreadsheet();
  if (ss.getSheetByName('指導者マスタ')) return;
  const sheet = ss.insertSheet('指導者マスタ');
  sheet.appendRow([
    'No', '氏名', 'クラブ名', '区分', '指導者種別', '時給',
    '金融機関', '支店名', '口座種別', '口座番号', 'メールアドレス', '登録日', 'PIN',
  ]);
  sheet.setFrozenRows(1);
}
