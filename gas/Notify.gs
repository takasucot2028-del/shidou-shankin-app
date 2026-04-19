/**
 * メール通知処理
 * CLAUDE.md の通知メール仕様に従い、各種メールを送信する
 * 送信結果は「通知ログ」シートに記録する
 */

// ========== 提出通知（指導者 → 事務局） ==========

/**
 * 月報提出時に事務局全員へメールを送る
 * Report.gs の submitReport から呼ばれる
 */
function notifyAdminOnSubmit(instructorName, clubName, year, month, submittedAt) {
  const adminEmails = getAdminEmails();
  if (adminEmails.length === 0) return;

  const subject = '【月報提出】' + clubName + ' ' + instructorName + '様より'
    + year + '年' + month + '月分が提出されました';

  const body = [
    '以下の月報が提出されました。',
    '',
    '指導者: ' + instructorName,
    'クラブ: ' + clubName,
    '対象月: ' + year + '年' + month + '月',
    '提出日時: ' + Utilities.formatDate(submittedAt, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss'),
    '',
    '管理画面からご確認ください。',
  ].join('\n');

  adminEmails.forEach(email => {
    sendMail(email, subject, body);
    logNotification(email, '提出通知', year + '年' + month + '月', '送信済');
  });
}

// ========== リマインドメール（事務局 → 指導者） ==========

/**
 * 未提出者へリマインドメールを送る
 * Trigger.gs のトリガーから呼ばれる
 * @param {number} daysBeforeDeadline - 締切までの残日数（メール文言に使用）
 */
function sendReminderEmails(daysBeforeDeadline) {
  const now = new Date();
  const year = now.getFullYear();
  // 当月5日が締切なので、当月の月報 = 前月分
  const targetMonth = now.getMonth(); // getMonth() は 0-indexed → 前月の月番号
  const targetYear = targetMonth === 0 ? year - 1 : year;
  const adjustedMonth = targetMonth === 0 ? 12 : targetMonth;

  const dashboard = getDashboard({
    year: String(targetYear),
    month: String(adjustedMonth),
  });

  if (!dashboard.success) return;

  // 締切は当月5日（提出対象は前月分）
  const deadlineMonth = now.getMonth() + 1; // 1-indexed 当月
  const deadline = year + '年' + deadlineMonth + '月5日';
  const daysText = daysBeforeDeadline ? '提出期限' + daysBeforeDeadline + '日前' : '提出期限当日';

  dashboard.unsubmittedList.forEach(item => {
    const instructor = findInstructor(item.instructorName);
    if (!instructor || !instructor.email) return;

    const subject = '【' + daysText + '】' + adjustedMonth + '月分 指導月報の提出をお願いします';
    const body = [
      instructor.name + ' 様',
      '',
      adjustedMonth + '月分の指導月報提出期限（' + deadline + '）が近づいています。',
      'まだ提出されていない場合は下記URLから提出をお願いします。',
      '',
      '提出URL: ' + getReportUrl(),
      '',
      '※期限を過ぎた場合は事務局にご連絡ください。',
      '',
      getOrgSignature(),
    ].join('\n');

    sendMail(instructor.email, subject, body);
    logNotification(instructor.email, 'リマインド', targetYear + '年' + adjustedMonth + '月', '送信済');
  });
}

// ========== 未提出者リスト通知（事務局向け） ==========

/**
 * 締切日に事務局へ未提出者リストを送る
 * Trigger.gs のトリガーから呼ばれる
 */
function notifyAdminUnsubmitted() {
  const adminEmails = getAdminEmails();
  if (adminEmails.length === 0) return;

  const now = new Date();
  const year = now.getFullYear();
  const targetMonth = now.getMonth();
  const targetYear = targetMonth === 0 ? year - 1 : year;
  const adjustedMonth = targetMonth === 0 ? 12 : targetMonth;

  const dashboard = getDashboard({
    year: String(targetYear),
    month: String(adjustedMonth),
  });

  if (!dashboard.success) return;

  const unsubmitted = dashboard.unsubmittedList;
  const subject = '【未提出者通知】' + targetYear + '年' + adjustedMonth + '月分 指導月報';

  const lines = [
    targetYear + '年' + adjustedMonth + '月分の月報締切日となりました。',
    '未提出者は以下の通りです（' + unsubmitted.length + '名）。',
    '',
  ];

  if (unsubmitted.length === 0) {
    lines.push('未提出者はいません。全員提出済みです。');
  } else {
    unsubmitted.forEach((item, idx) => {
      lines.push((idx + 1) + '. ' + item.instructorName + '（' + item.clubName + '）');
    });
  }

  lines.push('', '管理画面: ' + getAdminUrl());

  const body = lines.join('\n');

  adminEmails.forEach(email => {
    sendMail(email, subject, body);
    logNotification(email, '未提出者通知', targetYear + '年' + adjustedMonth + '月', '送信済');
  });
}

// ========== メール送信 ==========

function sendMail(to, subject, body) {
  try {
    GmailApp.sendEmail(to, subject, body);
  } catch (err) {
    console.error('メール送信失敗 to=' + to + ' err=' + err.message);
  }
}

// ========== 通知ログ記録 ==========

function logNotification(email, type, targetYm, status) {
  const sheet = getOrCreateNotifyLogSheet();
  sheet.appendRow([new Date(), email, type, targetYm, status]);
}

function getOrCreateNotifyLogSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName('通知ログ');
  if (!sheet) {
    sheet = ss.insertSheet('通知ログ');
    sheet.appendRow(['送信日時', '宛先メール', '通知種別', '対象年月', 'ステータス']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ========== 支払予定日計算 ==========

/**
 * 月報対象月に対する支払予定日を返す（翌月10日、土日祝なら前の平日）
 * @param {number} reportYear
 * @param {number} reportMonth
 * @return {Date}
 */
function calcPaymentDate(reportYear, reportMonth) {
  let y = reportYear, m = reportMonth + 1;
  if (m > 12) { m = 1; y++; }

  const holidays = getJPHolidaySet(y);
  const date = new Date(y, m - 1, 10);

  while (date.getDay() === 0 || date.getDay() === 6 || holidays.has(dateKey(date))) {
    date.setDate(date.getDate() - 1);
  }
  return date;
}

function dateKey(d) {
  return d.getFullYear() + '-' + (d.getMonth() + 1) + '-' + d.getDate();
}

/**
 * 指定年の日本の祝日Setを返す（dateKey形式）
 */
function getJPHolidaySet(year) {
  const set = new Set();
  const add = (m, d) => set.add(year + '-' + m + '-' + d);

  add(1, 1);
  add(2, 11);
  if (year >= 2020) add(2, 23);
  add(3, calcShunbunDay(year));
  add(4, 29);
  add(5, 3); add(5, 4); add(5, 5);
  add(8, 11);
  add(9, calcShubunDay(year));
  add(11, 3);
  add(11, 23);

  add(1,  nthMondayDay(year, 1,  2));
  add(7,  nthMondayDay(year, 7,  3));
  add(9,  nthMondayDay(year, 9,  3));
  add(10, nthMondayDay(year, 10, 2));

  // 振替休日
  const base = new Set(set);
  base.forEach(function(key) {
    const parts = key.split('-');
    const d = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
    if (d.getDay() === 0) {
      const sub = new Date(d);
      sub.setDate(sub.getDate() + 1);
      while (set.has(dateKey(sub))) sub.setDate(sub.getDate() + 1);
      set.add(dateKey(sub));
    }
  });

  return set;
}

function calcShunbunDay(year) {
  return Math.floor(20.8431 + 0.242194 * (year - 1980) - Math.floor((year - 1980) / 4));
}

function calcShubunDay(year) {
  return Math.floor(23.2488 + 0.242194 * (year - 1980) - Math.floor((year - 1980) / 4));
}

function nthMondayDay(year, month, n) {
  const d = new Date(year, month - 1, 1);
  let count = 0;
  while (true) {
    if (d.getDay() === 1) { count++; if (count === n) return d.getDate(); }
    d.setDate(d.getDate() + 1);
  }
}

// ========== URL・署名ヘルパー ==========

function getReportUrl() {
  return PropertiesService.getScriptProperties().getProperty('REPORT_URL') || 'https://takasucot2028-del.github.io/shidou-shankin-app/';
}

function getAdminUrl() {
  return PropertiesService.getScriptProperties().getProperty('ADMIN_URL') || 'https://takasucot2028-del.github.io/shidou-shankin-app/admin.html';
}

function getOrgSignature() {
  return PropertiesService.getScriptProperties().getProperty('ORG_SIGNATURE')
    || '〇〇町教育委員会 スポーツ・健康づくり担当';
}
