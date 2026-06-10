/**
 * Utils.gs (v4.0.4 料理名結合 ＋ 時刻ズレ・全角半角ゆらぎ完全吸収ソート版)
 * 役割：共通ロジック、ログ管理、ファイル操作、API通信、日付計算、一覧抽出
 */

function getCurrentSchoolYear() { return 2026; }
function getYearFromMonth(month) { return getCurrentSchoolYear(); }

function formatDate(date, format = "yyyy/MM/dd") {
  if (!date) return "";
  return Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), format);
}

function extractId(urlOrId) {
  if (!urlOrId) return "";
  const match = String(urlOrId).match(/[-\w]{25,}/);
  return match ? match[0] : urlOrId;
}

// ■ ログ管理
function writeLog(message, type = "info") {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheetName = (typeof CONFIG !== 'undefined' && CONFIG.SHEET_NAMES && CONFIG.SHEET_NAMES.LOG) ? CONFIG.SHEET_NAMES.LOG : "実行ログ";
  let sheet = ss.getSheetByName(logSheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(logSheetName);
    sheet.getRange(1, 1).setValue("📋 システム実行ログ").setFontWeight("bold");
    sheet.getRange(2, 1, 1, 3).setValues([["実行日時", "タイプ", "内容"]]).setBackground("#f3f3f3").setFontWeight("bold");
    sheet.setFrozenRows(2);
    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(3, 400);
  }

  try {
    sheet.insertRowBefore(3);
    const timestamp = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm:ss");
    const range = sheet.getRange(3, 1, 1, 3);
    range.setValues([[timestamp, type, message]]);
    
    if (type === "error") range.setFontColor("#FF0000").setFontWeight("bold");
    else if (type === "success") range.setFontColor("#008000");
    else range.setFontColor("#000000");
  } catch(e) { console.error("ログ記録失敗: " + e.message); }

  try {
    const cache = CacheService.getScriptCache();
    const key = "LOG_QUEUE";
    let queue = [];
    const cached = cache.get(key);
    if (cached) queue = JSON.parse(cached);
    queue.push(message);
    cache.put(key, JSON.stringify(queue), 21600);
  } catch (e) { }
}

// ■ ファイル操作・API通信系
function findFileByKeyword(folderId, keyword) {
  try {
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      if (file.getName().includes(keyword) && file.getMimeType() === MimeType.PDF) return file;
    }
  } catch(e) { writeLog(`❌ フォルダアクセスエラー: ${e.message}`, "error"); }
  return null;
}

function findAllFilesByKeyword(folderId, keyword) {
  const result = [];
  try {
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      if (file.getMimeType() === MimeType.PDF && file.getName().includes(keyword)) result.push(file);
    }
  } catch(e) { writeLog(`❌ フォルダ検索エラー: ${e.message}`, "error"); }
  return result;
}

function moveFileToFolder(file, targetFolderId) {
  const targetFolder = DriveApp.getFolderById(targetFolderId);
  file.moveTo(targetFolder);
}

function callGeminiWithRetry(apiKey, blob, prompt, maxRetries = 3) {
  const model = (typeof CONFIG !== 'undefined' && CONFIG.MODEL_NAME) ? CONFIG.MODEL_NAME : 'gemini-1.5-flash';
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;
  
  const payload = {
    contents: [{ parts: [{ text: prompt }, { inline_data: { mime_type: "application/pdf", data: Utilities.base64Encode(blob.getBytes()) } }] }]
  };
  const options = { method: "post", contentType: "application/json", payload: JSON.stringify(payload), muteHttpExceptions: true };

  for (let i = 0; i < maxRetries; i++) {
    try {
      const response = UrlFetchApp.fetch(url, options);
      const resCode = response.getResponseCode();
      const resText = response.getContentText();
      const json = JSON.parse(resText);

      if (resCode === 200) {
        if (!json.candidates || !json.candidates[0].content) throw new Error("応答が空");
        return json.candidates[0].content.parts[0].text;
      }
      if (resText.includes("quota") || resCode === 429) {
        writeLog(`⏳ API制限待機中。60秒後に再開（リトライ ${i+1}/${maxRetries}）`, "info");
        Utilities.sleep(60000);
        continue;
      }
      const errorDetail = json.error ? json.error.message : "不明なエラー";
      writeLog(`❌ APIエラー(${i+1}/${maxRetries}): ${errorDetail}`, "error");
      Utilities.sleep(5000);
    } catch (e) {
      writeLog(`⚠️ システムエラー(${i+1}/${maxRetries}): ${e.message}`, "warn");
      Utilities.sleep(5000);
    }
  }
  throw new Error(`API呼び出しに失敗しました。`);
}

function parseGeminiOutput(text, minColumns = 1) {
  if (!text) return [];
  const lines = text.replace(/```/g, "").trim().split(/\r?\n/);
  const results = [];
  lines.forEach(line => {
    if (!line.trim()) return;
    const parts = line.split("|").map(s => s.trim());
    if (parts.length >= minColumns) results.push(parts);
  });
  return results;
}

function getLogQueue() {
  const cache = CacheService.getScriptCache();
  const key = "LOG_QUEUE";
  const cached = cache.get(key);
  if (cached) { cache.remove(key); try { return JSON.parse(cached); } catch (e) { return []; } }
  return [];
}

