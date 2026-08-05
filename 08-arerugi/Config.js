/**
 * Config.gs
 * VERSION: v4.6.0
 * 役割：システム全体の設定値管理、APIキー・IDの取得、判定無視ワードの動的取得
 */
const CONFIG = {
  VERSION: 'v4.6.0',
  MODEL_NAME: 'gemini-flash-latest',

  // 対象年度（令和8年は2026年）
  TARGET_YEAR: 2026,

  // シート名の定義
  SHEET_NAMES: {
    CONFIG: '基礎データ',
    MASTER: '献立マスタ',
    MAIN: 'Main',
    DICT: 'アレルゲン辞書',
    VERIFY: '確認用一覧',
    DASHBOARD: 'ダッシュボード',
    LOG: '実行ログ',
    PERSONAL: '個人アレルゲン', 
    ALERT: '記入漏れアラート'   
  },

  // フォルダ名定義
  FOLDERS: {
    BACKUP: "📦_バックアップ保存箱"
  },

  // データ開始行
  START_ROW: 3,

  // ダッシュボード設定
  DASHBOARD: {
    DATE_CELL: "B2",
    TARGET_COUNT_CELL: "E2",
    REFRESH_RANGE: "B8:G30"
  },

  // ==========================================
  // 🚨 設定読み取り機能
  // ==========================================

  // Bot（カレンダー同期・通知）がONになっているか確認する
  isBotActive: function() {
    try {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAMES.CONFIG);
      const val = sheet.getRange("B8").getValue();
      
      if (val === true) return true;
      if (String(val).toUpperCase() === "TRUE") return true;
      if (String(val).toUpperCase() === "ON") return true;
      if (val === 1) return true;
      
      return false; 
    } catch(e) { 
      return false; 
    }
  },

  // カレンダーIDを取得する
  getCalendarId: function() {
    try {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAMES.CONFIG);
      return String(sheet.getRange("B5").getValue() || "").trim();
    } catch(e) { 
      return ""; 
    }
  },

  // チャット通知用のWebhook URLを取得する
  getWebhookUrl: function() {
    try {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAMES.CONFIG);
      return String(sheet.getRange("B7").getValue() || "").trim();
    } catch(e) { 
      return ""; 
    }
  },

  // 🌟【新規追加】記入漏れアラート用の判定無視ワード（B10セル）を動的に取得・配列化する
  getIgnoreWords: function() {
    try {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAMES.CONFIG);
      const val = sheet.getRange("B10").getValue();
      if (!val) return [];
      // 読点、カンマ、スペース、改行のいずれかで区切られたキーワードを綺麗に配列に分解
      return String(val).split(/[、，,\s\n\r]+/).map(k => k.trim()).filter(Boolean);
    } catch(e) { 
      return []; 
    }
  },

  // 🌟【新規追加】APIキーをB2セルから取得し、内部（スクリプトプロパティ）に格納・管理する
  getApiKey: function() {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      // 未確定のセル編集を強制確定
      SpreadsheetApp.flush();
      const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.CONFIG);
      const cell = sheet.getRange("B2");
      let val = String(cell.getValue() || "").trim();
      
      const propKey = "GEMINI_API_KEY";
      const props = PropertiesService.getScriptProperties();
      
      // 新しいAPIキーが入力された場合（非表示マスクなどの文字列以外で値がある場合）
      if (val && val !== "（設定済み）" && val !== "（格納済み）" && !val.includes("非表示")) {
        props.setProperty(propKey, val);
        cell.setValue("（設定済み）");
        SpreadsheetApp.flush(); // 即時セル表示を更新
        ss.toast("B2セルの新しいAPIキーを内部（スクリプトプロパティ）に格納しました。", "🔑 APIキー更新", 5);
        if (typeof writeLog === 'function') {
          writeLog("🔑 B2セルのAPIキーをスクリプトプロパティに格納しました。", "success");
        }
        return val;
      }
      
      // 内部に保存されているキーを返す
      const savedKey = props.getProperty(propKey);
      if (savedKey) {
        return savedKey;
      }
      
      // 内部に保存されていない場合はセルに入っている値をそのまま返す
      return val;
    } catch(e) {
      if (typeof writeLog === 'function') {
        writeLog("⚠️ APIキーの取得中にエラーが発生しました: " + e.message, "error");
      }
      return "";
    }
  }
};