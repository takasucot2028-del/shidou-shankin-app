/**
 * トリガー設定
 * CLAUDE.md のトリガー仕様に従い、時間ベーストリガーを登録・削除する
 *
 * 【初回セットアップ手順】
 * 1. GAS エディタで setupAllTriggers() を手動実行してください
 * 2. 既存トリガーを再設定する場合は deleteAllTriggers() → setupAllTriggers() の順で実行
 */

// ========== トリガー一括登録 ==========

/**
 * 全トリガーを登録する（初回セットアップ用）
 */
function setupAllTriggers() {
  deleteAllTriggers();

  // 毎月1日 00:00 → リマインド開始準備（翌月5日締切の通知スケジュール起動）
  createMonthlyTrigger('onFirstOfMonth', 1, 0);

  // 毎月3日 09:00 → 未提出者へリマインドメール（締切3日前相当）
  createMonthlyTrigger('onThirdOfMonth', 3, 9);

  // 毎月4日 09:00 → 未提出者へ最終リマインドメール
  createMonthlyTrigger('onFourthOfMonth', 4, 9);

  // 毎月5日 18:00 → 事務局へ未提出者リスト通知
  createMonthlyTrigger('onFifthOfMonth', 5, 18);

  console.log('全トリガーを登録しました');
}

// ========== 月次トリガーハンドラ ==========

/** 毎月1日 00:00 */
function onFirstOfMonth() {
  // 当月5日締切のリマインド送信開始（3・4日のトリガーで実際に送信）
  // ここでは未提出者の事前確認のみ（必要に応じて拡張）
  console.log('月次処理開始: ' + new Date());
}

/** 毎月3日 09:00 — 締切3日前リマインド */
function onThirdOfMonth() {
  sendReminderEmails(3);
}

/** 毎月4日 09:00 — 締切前日リマインド（最終） */
function onFourthOfMonth() {
  sendReminderEmails(1);
}

/** 毎月5日 18:00 — 締切後、事務局へ未提出者通知 */
function onFifthOfMonth() {
  notifyAdminUnsubmitted();
}

// ========== トリガー作成ヘルパー ==========

/**
 * 毎月指定日・時刻に実行されるトリガーを作成する
 * GAS の月次トリガーは「毎月X日」指定が直接できないため、
 * 日次トリガーを使い、ハンドラ側で日付を判定する
 */
function createMonthlyTrigger(funcName, day, hour) {
  ScriptApp.newTrigger(funcName)
    .timeBased()
    .onMonthDay(day)
    .atHour(hour)
    .create();
}

// ========== 手動実行用ヘルパー ==========

/**
 * 手動実行・緊急送信用リマインド
 * GASエディタから直接実行してテストや緊急対応に使用する
 * @param {number} daysBeforeDeadline - 締切までの残日数（省略時は2）
 */
function sendReminder(daysBeforeDeadline) {
  const days = (daysBeforeDeadline != null) ? daysBeforeDeadline : 2;
  sendReminderEmails(days);
  console.log('リマインドメール送信完了（締切' + days + '日前）');
}

/**
 * 手動実行・緊急送信用 事務局未提出者通知
 * GASエディタから直接実行してテストや緊急対応に使用する
 */
function sendAdminUnsubmittedNotice() {
  notifyAdminUnsubmitted();
  console.log('事務局への未提出者通知送信完了');
}

// ========== トリガー削除 ==========

/**
 * プロジェクトに登録された全トリガーを削除する
 */
function deleteAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  console.log(triggers.length + ' 件のトリガーを削除しました');
}

// ========== トリガー一覧確認 ==========

/**
 * 登録済みトリガーをログに出力して確認する
 */
function listTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  if (triggers.length === 0) {
    console.log('登録されているトリガーはありません');
    return;
  }
  triggers.forEach(t => {
    console.log(
      '関数: ' + t.getHandlerFunction() +
      ' | 種別: ' + t.getEventType() +
      ' | ソース: ' + t.getTriggerSource()
    );
  });
}

// ========== 初回セットアップ（スプレッドシート初期化込み） ==========

/**
 * GAS を初めてデプロイした直後に実行する
 * 全シートの初期化 + トリガー登録を一括で行う
 */
function initialSetup() {
  initMasterSheet();       // Master.gs
  getOrCreateFeeSheet();   // Fee.gs
  getOrCreateNotifyLogSheet(); // Notify.gs
  setupAllTriggers();
  console.log('初期セットアップ完了');
}
