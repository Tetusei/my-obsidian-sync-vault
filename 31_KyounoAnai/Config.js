/**
 * Config.gs
 * VERSION: v48.4.0 (2026年度 運用安定版・トリガー管理追加)
 */

const CONFIG = {
  VERSION: "v48.4.0", // インクリメント
  SHEET_NAME_CONFIG: '基礎データ',
  SHEET_NAME_BOT: '連絡・日課',
  SCHEDULE_START_ROW: 10,
  MODEL_NAME: "gemini-flash-latest",
  BACKUP_FOLDER_NAME: "📦_バックアップ保存箱",
  DEFAULT_YEAR: 2026, 
  FISCAL_YEAR: 2026
};

/**
 * 基礎データシートから設定を取得
 */
function getSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME_CONFIG);
  if (!sheet) {
    SpreadsheetApp.getUi().alert("「基礎データ」シートが見つかりません。");
    throw new Error("Sheet not found: " + CONFIG.SHEET_NAME_CONFIG);
  }

  const values = sheet.getRange("B1:C12").getValues();
  const apiKeys = [values[0][0], values[1][0], values[2][0]].filter(k => k && k.toString().trim() !== "");

  if (apiKeys.length === 0) {
    throw new Error("APIキーが設定されていません。");
  }
 
  return {
    apiKeys: apiKeys,
    eventCalId: cleanCalendarId(values[3][0]),  
    menuCalId: cleanCalendarId(values[4][0]),   
    pdfFolderId: extractFolderId(values[5][0]), 
    webhookUrl: values[6][0], 
    isBotEnabled: values[7][0] === "ON", 
    countName: values[8][0], 
    countTargetDate: values[9][0], 
    isAiEnabled: values[10][0] === "ON", 
    areaName: values[11][0], 
    areaCode: values[11][1]
  };
}

function cleanCalendarId(id) {
  if (!id || typeof id !== 'string') return id;
  const match = id.match(/src=([^&]+)/);
  return match ? decodeURIComponent(match[1]) : id.trim();
}

function extractFolderId(url) {
  if (!url || typeof url !== 'string') return url;
  const match = url.match(/folders\/([^\/\?]+)/);
  return match ? match[1] : url;
}