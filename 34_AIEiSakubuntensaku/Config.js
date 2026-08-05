/**
 * ファイル名: Config.gs
 * 役割: 定数・設定管理
 * バージョン: v4.0.0 (F1セル・22時自動締め切り対応 完全版)
 */

const Config = {
  VERSION: 'v4.0.0',
  MODEL_NAME: 'gemini-flash-latest', 
  TARGET_YEAR: 2026,
  
  BACKUP_FOLDER_NAME: '📦_バックアップ保存箱',
  
  API_KEY_SHEET: '基礎データ', 
  API_KEY_RANGE: 'B1:B3',      
  LINK_SWITCH_CELL: 'B4',      
  MODE_SWITCH_CELL: 'B5',      
  PROMPT_CELL: 'B6',           
  ROSTER_SHEET: '名簿',
  REPORT_SHEET: '📊_分析レポート', 
  MAIN_SHEET: '添削',        
  QUESTION_CELL: 'B1',       
  PROGRESS_CELL: 'C1',       
  DEADLINE_CELL: 'F1',       // ★E1からF1に変更しました
  
  // --- データの列番号・行番号設定 ---
  DATA_START_ROW: 3,        // データが始まる行番号（3行目）
  COL_NUM: 1,               // A列: 番号
  COL_STUDENT_NAME: 2,      // B列: 氏名
  COL_STUDENT: 3,           // C列: 生徒の英文
  COL_FIX: 4,               // D列: 添削結果
  COL_EXPLAIN: 5,           // E列: 先生からの解説
  COL_JA: 6,                // F列: 和訳
  COL_SCORE: 7,             // G列: AI評価スコア
  COL_TIMESTAMP: 8,         // H列: 提出日時
  COL_MAIL_STATUS: 9        // I列: メール送信状況
};

function getApiKeys() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(Config.API_KEY_SHEET);
  if (!sheet) throw new Error(`「${Config.API_KEY_SHEET}」シートが見つかりません。`);
  
  const values = sheet.getRange(Config.API_KEY_RANGE).getValues();
  const keys = [];
  const docProperties = PropertiesService.getDocumentProperties();
  let changed = false;

  for (let i = 0; i < values.length; i++) {
    const val = values[i][0] ? values[i][0].toString().trim() : "";
    if (val === "設定済み") {
      const storedKey = docProperties.getProperty('API_KEY_' + i);
      if (storedKey && storedKey.trim() !== '') {
        keys.push(storedKey.trim());
      }
    } else if (val !== "") {
      // 生のキーが入っている場合は、プロパティに退避し「設定済み」にする（フォールバック）
      docProperties.setProperty('API_KEY_' + i, val);
      keys.push(val);
      sheet.getRange(i + 1, 2).setValue("設定済み"); // B1は(1,2), B2は(2,2), B3は(3,2)
      changed = true;
    } else {
      // 空欄の場合は念のためプロパティも削除
      const storedKey = docProperties.getProperty('API_KEY_' + i);
      if (storedKey) {
        docProperties.deleteProperty('API_KEY_' + i);
      }
    }
  }

  if (changed) {
    SpreadsheetApp.flush();
  }

  if (keys.length === 0) throw new Error(`APIキーが設定されていません。`);
  return keys;
}

function getCustomPrompt() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(Config.API_KEY_SHEET);
  if (!sheet) return "";
  const promptText = sheet.getRange(Config.PROMPT_CELL).getValue();
  return promptText ? promptText.toString() : "";
}

function getActionMode() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(Config.API_KEY_SHEET);
  if (!sheet) return "";
  const mode = sheet.getRange(Config.MODE_SWITCH_CELL).getValue();
  return mode ? mode.toString() : "";
}