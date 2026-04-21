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
  const allData = sheet.getDataRange().getValues();
  const headers = allData[0];
  const pinIdx  = headers.indexOf('PIN'); // -1 の場合は列が存在しない

  for (let i = 1; i < allData.length; i++) {
    if (allData[i][1] === body.name) {
      const existingPin = pinIdx !== -1 ? String(allData[i][pinIdx] || '') : '';
      const newPin = (body.pin !== undefined && body.pin !== '') ? String(body.pin) : existingPin;

      const newRow = [
        allData[i][0],
        body.name,
        body.clubName,
        body.rateType,
        body.instructorType,
        body.hourlyRate,
        body.bank           || '',
        body.branch         || '',
        body.accountType    || '',
        body.accountNumber  || '',
        body.email          || '',
        allData[i][11],       // 登録日は変えない
        newPin,               // PIN（M列 = インデックス12）
      ];
      sheet.getRange(i + 1, 1, 1, newRow.length).setValues([newRow]);
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
  const allData = sheet.getDataRange().getValues();
  const headers = allData[0];
  const pinIdx  = headers.indexOf('PIN');

  if (pinIdx === -1) {
    return { success: false, error: 'PIN列がシートに存在しません。事務局にお問い合わせください' };
  }

  const row = allData.slice(1).find(r => r[1] === body.name);
  if (!row) {
    return { success: false, error: 'PINが正しくありません' };
  }
  const storedPin = String(row[pinIdx] || '');
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

// ========== PIN列マイグレーション ==========

/**
 * 既存の「指導者マスタ」シートにPIN列（M列）がなければ追加する
 * PIN列追加前にシートを作成した場合に1回だけ手動実行する
 */
function setupPinColumn() {
  const sheet   = getSheet('指導者マスタ');
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  if (headers.indexOf('PIN') !== -1) {
    Logger.log('PIN列は既に存在します（列番号: ' + (headers.indexOf('PIN') + 1) + '）');
    return;
  }

  // M列（13列目）にPINヘッダーを追加
  const pinCol = 13;
  sheet.getRange(1, pinCol).setValue('PIN');

  // 既存データ行のPINセルを空文字で初期化
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, pinCol, lastRow - 1, 1).setValue('');
  }

  Logger.log('PIN列をM列（13列目）に追加しました（データ行数: ' + (lastRow - 1) + '）');
}

/**
 * Apps Scriptエディタから直接実行するためのラッパー
 * 「指導者マスタ」シートにPIN列がない場合に1回だけ実行してください
 */
function runSetupPinColumn() {
  setupPinColumn();
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

/**
 * 初期指導者データを一括投入する
 * 「指導者マスタ」シートが空（ヘッダーのみ）の状態で手動実行
 * 既にデータがある場合は何もしない
 */
function seedInstructors() {
  const sheet = getSheet('指導者マスタ');
  if (sheet.getLastRow() > 1) {
    Logger.log('既にデータが存在するためスキップしました（行数: ' + (sheet.getLastRow() - 1) + '）');
    return;
  }

  const RATE = { 'メイン': 1600, 'サブ': 1100 };
  const now = new Date();

  // [No, 氏名, クラブ名, 区分, 指導者種別]
  const instructors = [
    [1,  '髙畠　奈々恵',   'NexusBC',              'メイン', '一般'],
    [2,  '髙畠　茂樹',     'NexusBC',              'サブ',   '一般'],
    [3,  '中塩　春菜',     'NexusBC',              'サブ',   '一般'],
    [4,  '井上　操',       'バドミントンクラブ',    'サブ',   '一般'],
    [5,  '遠藤　和馬',     'バドミントンクラブ',    'サブ',   '教員'],
    [6,  '菅野　聖一',     'REDWOLVES男子',         'メイン', '一般'],
    [7,  '山本　晃司',     'REDWOLVES男子',         'サブ',   '教員'],
    [8,  '小野寺　隼也',   'REDWOLVES男子',         'メイン', '一般'],
    [9,  '二階堂　孝',     'REDWOLVES男子',         'サブ',   '一般'],
    [10, '中井　啓太',     'REDWOLVES男子',         'サブ',   '一般'],
    [11, '高瀬　雅裕',     'REDWOLVES女子',         'メイン', '一般'],
    [12, '西側　莉',       'REDWOLVES女子',         'メイン', '一般'],
    [13, '西間　沙織',     'REDWOLVES女子',         'サブ',   '一般'],
    [14, '伊藤　睦',       'TakasuXC',              'メイン', '一般'],
    [15, '上坂　篤',       'TakasuXC',              'メイン', '一般'],
    [16, '高橋　朋亮',     '鷹高館',                'メイン', '教員'],
    [17, '後木　主税',     '鷹高館',                'サブ',   '一般'],
    [18, '長田　良信',     '鷹高館',                'サブ',   '一般'],
    [19, '杉森　貴之',     'たかすテニスクラブ',    'メイン', '教員'],
    [20, '森木　貴仁',     'たかすテニスクラブ',    'サブ',   '一般'],
    [21, '松本　圭代',     'たかすテニスクラブ',    'サブ',   '教員'],
    [22, '館田　梨奈',     '男子バレーボールクラブ','メイン', '一般'],
    [23, '辰巳　実莉',     '男子バレーボールクラブ','サブ',   '一般'],
    [24, '田中　慎二',     '女子バレーボールクラブ','メイン', '教員'],
    [25, '堀　朱音',       '女子バレーボールクラブ','サブ',   '教員'],
    [26, '上原　丈典',     '男子バレーボールクラブ','サブ',   '教員'],
    [27, '宮下　隆太郎',   '野球クラブ',            'メイン', '教員'],
    [28, '山崎　真由美',   '野球クラブ',            'サブ',   '教員'],
    [29, '吉峯　浩二郎',   '野球クラブ',            'サブ',   '教員'],
    [30, '長田　頼子',     '鷹栖吹奏楽クラブ',      'メイン', '教員'],
    [31, '今野　佳代',     '鷹栖吹奏楽クラブ',      'サブ',   '教員'],
    [32, '甲斐先生',       'たかすスクールバンド',  'メイン', '教員'],
    [33, '安達　一幸',     'マルチスポーツクラブ',  'サブ',   '一般'],
  ];

  const rows = instructors.map(([no, name, club, rateType, instrType]) => [
    no, name, club, rateType, instrType,
    RATE[rateType],
    '', '', '', '', '', // 金融機関・支店・口座種別・口座番号・メール（後で個別設定）
    now,
    '', // PIN（後で個別設定）
  ]);

  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  Logger.log('指導者マスタに ' + rows.length + ' 件を登録しました');
}
