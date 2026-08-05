/**
 * Ai_Pdf.gs
 * VERSION: v49.3 (重複排除・完全整合版)
 * AI連携、PDF解析、および補助ボタン（前日・翌日）ロジック
 */

/**
 * 1. Gemini API呼び出し
 */
function callGeminiWithRetry(model, promptText, settings, pdfBlob = null) {
  const keys = Array.isArray(settings.apiKeys) ? settings.apiKeys : [settings.apiKeys];
  for (let key of keys) {
    if (!key) continue;
    try {
      const parts = [{ text: promptText }];
      if (pdfBlob) {
        let base64Data = pdfBlob;
        if (typeof pdfBlob !== 'string') { 
           base64Data = Utilities.base64Encode(pdfBlob.getBytes());
        }
        parts.push({ inline_data: { mime_type: "application/pdf", data: base64Data } });
      }
      const payload = { contents: [{ parts: parts }] };
      const options = {
        method: 'post', contentType: 'application/json', payload: JSON.stringify(payload),
        muteHttpExceptions: true, timeoutSeconds: 300
      };
      const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${key}`;
      const res = UrlFetchApp.fetch(url, options);
      if (res.getResponseCode() === 200) {
        const json = JSON.parse(res.getContentText());
        if (json.candidates && json.candidates[0].content) {
          return json.candidates[0].content.parts[0].text;
        }
      }
    } catch (e) { console.warn("APIリトライ中: " + e.message); }
  }
  return null;
}

/**
 * 2. ファイルリスト取得
 */
function getClientFileList() {
  const settings = getSettings();
  try {
    if (!settings.pdfFolderId) return [{id:"", name:"⚠️ 基礎データ(B6)のフォルダIDが空です"}];
    const folder = DriveApp.getFolderById(settings.pdfFolderId);
    const files = folder.getFilesByType(MimeType.PDF);
    const list = [];
    while(files.hasNext()){ let f = files.next(); list.push({id:f.getId(), name:f.getName()}); }
    if (list.length === 0) return [{id:"", name:"(フォルダにPDFがありません)"}];
    return list;
  } catch(e) { return [{id:"", name:"❌ フォルダID設定を確認してください"}]; }
}

/**
 * 3. PDF解析
 */
function processPdf(fileId, mode) {
  const settings = getSettings();
  try {
    const file = DriveApp.getFileById(fileId);
    const blobStr = Utilities.base64Encode(file.getBlob().getBytes());
    
    let prompt = "";
    if (mode === 'Event') {
      prompt = `あなたは学校の事務支援AIです。添付の「行事予定PDF」を解析し、2026年（令和8年）のデータを以下のJSONで出力してください。\n{"Events":[{"StartDate":"YYYY-MM-DD","StartTime":"HH:MM","Subject":"行事名","Note":""}]}\n※Markdown記法不要`;
    } else {
      prompt = `あなたは学校の給食担当AIです。添付の「献立表PDF」を解析し、2026年（令和8年）のデータを以下のJSONで出力してください。\n{"Events":[{"StartDate":"YYYY-MM-DD","Subject":"メニュー内容"}]}\n※Markdown記法不要`;
    }

    const responseText = callGeminiWithRetry(CONFIG.MODEL_NAME, prompt, settings, blobStr);
    if (!responseText) throw new Error("AI応答なし");
    
    const jsonMatch = responseText.match(/\{[\s\S]*\}/);
    if (!jsonMatch) throw new Error("JSON抽出失敗");
    let json = JSON.parse(jsonMatch[0]);

    const safeName = file.getName().substring(0, 15).trim();
    const sheetName = `${safeName}_${mode === 'Event' ? '行事' : '給食'}`;
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(sheetName);
    if (sheet) sheet.clear(); else sheet = ss.insertSheet(sheetName);

    const headers = [["Chk", "開始", "件名", "終了", "終日", "説明", "状態"]];
    sheet.getRange(1, 1, 1, 7).setValues(headers).setBackground("#cfe2f3").setFontWeight("bold");

    const rows = json.Events.map(e => {
      const startDate = new Date(e.StartDate);
      if (isNaN(startDate.getTime())) return null;
      let isAllDay = true;
      if (mode === 'Event' && e.StartTime) isAllDay = false;
      return [true, startDate, e.Subject, startDate, isAllDay, e.Note || (mode === 'Menu'?'給食':''), '未'];
    }).filter(r => r !== null);

    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, 7).setValues(rows);
      sheet.getRange(2, 1, rows.length, 1).insertCheckboxes();
      sheet.getRange(2, 2, rows.length, 1).setNumberFormat("yyyy/mm/dd(ddd)");
      sheet.getRange(2, 4, rows.length, 1).setNumberFormat("yyyy/mm/dd(ddd)");
    }
    return { msg: `✅ 作成完了 (${rows.length}件)` };
  } catch(e) { return { msg: "❌ エラー: " + e.message, error: true }; }
}

// ==========================================
//  4. 補助ボタン用
// ==========================================

function movePrevDay() { shiftDate(-1); }
function moveNextDay() { shiftDate(1); }

function shiftDate(offset) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME_BOT);
  if (!sheet) return;
  const dateCell = sheet.getRange("A2");
  const date = dateCell.getValue();
  if (date instanceof Date) {
    date.setDate(date.getDate() + offset);
    dateCell.setValue(date);
    importCalendarToBotSheet(); // CalendarManager.gs
    if (typeof setDetailedWeather === 'function') setDetailedWeather(sheet, date);
  }
}

/**
 * 「行を追加」ボタン用
 */
function addNewManualRow() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();
  if (sheet.getName() === CONFIG.SHEET_NAME_BOT) {
    sheet.appendRow([true, "", "", "", "", "", "未"]);
    const newRow = lastRow + 1;
    sheet.getRange(newRow, 1).insertCheckboxes();
    sheet.getRange(newRow, 2).setNumberFormat("@");
  } else {
    sheet.appendRow([true, "", "", "", false, "", "未"]);
    const newRow = lastRow + 1;
    sheet.getRange(newRow, 1).insertCheckboxes();
    sheet.getRange(newRow, 2).setNumberFormat("yyyy/mm/dd(ddd)");
  }
}