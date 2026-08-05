/**
 * ============================================================================
 * ファイル名: Config.gs
 * 概要: システム全体の環境設定、定数、バージョン情報を一元管理するファイル。
 * ============================================================================
 */

const Config = {
  // 🚨 システムバージョン管理
  VERSION: 'v1.2.0',

  // 📅 処理対象年度の設定（2026年 / 令和8年度基準）
  TARGET_YEAR: 2026,

  // 📁 バックアップ設定
  BACKUP_FOLDER_NAME: '📦_バックアップ保存箱',

  // 🎨 デザイン設定（灰色以外の行ごと色分け用カラーコード）
  // 先生方が画面を長時間見ても目が疲れにくい、淡いパステルミントグリーンを採用
  COLORS: {
    EVEN: '#FFFFFF', // 奇数行（白色）
    ODD: '#F0F9F4'   // 偶数行（淡いミントグリーン）
  },

  // 📄 シート名設定
  SHEET_NAME_BASE_DATA: '基礎データ',
  SHEET_NAME_ROSTER: '名簿',
  SHEET_NAME_RECORD_PREFIX: 'R8年度',

  // 📍 セル位置設定
  WEBHOOK_CELL: 'B1',
  TARGET_PERSON_CELL: 'B3',
  PROPERTY_KEY_GEMINI_API: 'GEMINI_API_KEY',

  // 🔍 Gmail検索クエリ設定
  SEARCH_QUERY: 'subject:"見守りフィルター" is:unread',

  // 🤖 Geminiモデル設定
  GEMINI_MODEL: 'gemini-2.5-flash',

  // ⚙️ トリガー設定
  TRIGGER_FUNCTION_MAIL: 'checkGmailAndNotify',
  TRIGGER_INTERVAL_MINUTES: 10
};

// 互換性維持のためのエイリアス定義
const SystemConfig = Config;