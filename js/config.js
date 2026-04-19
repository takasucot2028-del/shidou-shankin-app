/**
 * GAS URL・アプリ共通設定
 * 本番環境に合わせて GAS_URL を変更してください
 */
const Config = {
  // Google Apps Script Web App URL
  // デプロイ後に取得したURLを設定する
  GAS_URL: 'https://script.google.com/macros/s/AKfycbzkgcqHJLn5fMX7mTxsrG1Z58yAt-5sD3CwtuSszhPHFZNGr6Au2nwFwG_MapbO4vgurA/exec',

  // 時給設定（円）
  HOURLY_RATE: {
    MAIN: 1600,
    SUB: 1100,
  },

  // 源泉徴収率
  WITHHOLDING_TAX_RATE: 0.1021,

  // 区分別上限時間（時間）
  MAX_HOURS: {
    '平日': 2,
    '休日': 3,
    '長期休暇': 3,
    '大会引率': 4,
  },

  // 教員の勤務終了時刻（分換算）
  TEACHER_WORK_END_MINUTES: 16 * 60 + 45, // 16:45

  // 月報提出締切日
  SUBMISSION_DEADLINE_DAY: 5,

  // ステータス定義
  STATUS: {
    DRAFT: '下書き',
    SUBMITTED: '提出済',
  },

  // 区分
  CATEGORY: {
    WEEKDAY: '平日',
    HOLIDAY: '休日',
    LONG_VACATION: '長期休暇',
    TOURNAMENT: '大会引率',
  },

  // 時給区分
  RATE_TYPE: {
    MAIN: 'メイン',
    SUB: 'サブ',
  },

  // 指導者種別
  INSTRUCTOR_TYPE: {
    TEACHER: '教員',
    GENERAL: '一般',
  },
};
