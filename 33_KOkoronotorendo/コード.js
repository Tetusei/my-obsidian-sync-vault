/**
 * 【No.53】心の健康観察システム：リンク切れ（#REF!）完全防止版
 * ・PDF作成時に「計算式」を「値」に変換して固定する処理を追加
 * ・これにより別シート参照(XLOOKUP等)による#REF!エラーを回避します
 * ・印刷範囲は A1:J54 に固定
 */

// ▼▼ 設定エリア ▼▼
const CONFIG_SHEET_NAME = '基礎データ';
const FORM_SHEET_NAME = 'フォームの回答 1'; 
const SURVEY_SHEET_NAME = '生活実態調査'; 
const PERSONAL_SHEET_NAME = '個人別'; 
const API_KEY_CELL = 'B1';
const START_ROW_STUDENT = 3; 
// ▲▲ 設定エリア終了 ▲▲

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('★健康観察') 
    .addItem('【手動実行】日記分析＆集計', 'analyzeAndAggregate')
    .addSeparator()
    .addItem('【手動実行】個人別データの更新', 'loadPersonalDataManual') 
    .addItem('【保存】表示中の生徒をPDF保存', 'createPdf') 
    .addSeparator()
    .addItem('【一括】全生徒のPDFを作成・結合', 'createMergedPdfFromSheets') 
    .addSeparator()
    .addItem('【分析】生活実態×日記 クロス分析', 'runCrossAnalysis') 
    .addToUi();
}

function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.source.getActiveSheet();
  const range = e.range;
  const row = range.getRow();
  const col = range.getColumn();
  const sheetName = sheet.getName();

  // 1. 個人別シートの切り替え処理
  if (sheetName === PERSONAL_SHEET_NAME) {
    if ((row === 1 && col === 2) || (row === 3 && col === 1)) {
      loadPersonalData(true);
    }
  }

  // 2. 基礎データシート의 APIキー安全保管処理
  if (sheetName === CONFIG_SHEET_NAME && row === 1 && col === 2) {
    const val = range.getValue().toString().trim();
    if (val !== "" && val !== "設定済み") {
      try {
        PropertiesService.getScriptProperties().setProperty("GEMINI_API_KEY", val);
        range.setValue("設定済み");
        SpreadsheetApp.flush();
        e.source.toast("APIキーを安全に保存しました。", "保存完了", 5);
      } catch (err) {
        e.source.toast("APIキーの保存に失敗しました: " + err.message, "エラー", 10);
      }
    } else if (val === "") {
      try {
        PropertiesService.getScriptProperties().deleteProperty("GEMINI_API_KEY");
        e.source.toast("APIキーを消去しました。", "消去完了", 5);
      } catch (err) {
        e.source.toast("APIキーの消去に失敗しました: " + err.message, "エラー", 10);
      }
    }
  }
}

// ==========================================
// ■ 機能1：日々の分析
// ==========================================
function analyzeAndAggregate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formSheet = ss.getSheetByName(FORM_SHEET_NAME);
  const configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!formSheet || !configSheet) return;

  const apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  if (!apiKey) {
    SpreadsheetApp.getUi().alert("エラー", "APIキーが設定されていません。基礎データシートのB1セルに入力してください。", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  const studentMap = getStudentMap(configSheet);
  const studentList = getStudentList(configSheet);
  const lastRowForm = formSheet.getLastRow();
  if (lastRowForm < 2) return;

  const formRange = formSheet.getRange(2, 1, lastRowForm - 1, 5);
  const formValues = formRange.getValues();
  let processCount = 0;

  for (let i = 0; i < formValues.length; i++) {
    const row = formValues[i];
    const email = row[1].toString().trim().toLowerCase(); 
    const score = row[2]; 
    const diary = row[3]; 
    const status = row[4]; 
    if (status !== "済" && diary && diary.toString() !== "") {
      const student = studentMap[email];
      if (!student) continue; 
      try {
        const aiResult = callGeminiForMental(apiKey, student.name, score, diary);
        writeToMonthlySheet(ss, studentList, new Date(row[0]), student, score, diary, aiResult);
        formValues[i][4] = "済";
        processCount++;
        Utilities.sleep(500); 
      } catch (e) {
        console.error(e.message);
      }
    }
  }
  if (processCount > 0) {
    formSheet.getRange(2, 5, formValues.length, 1).setValues(formValues.map(r => [r[4]]));
  }
}

// ==========================================
// ■ 機能2：データ読み込み
// ==========================================
function loadPersonalDataManual() {
  loadPersonalData(false);
}

function loadPersonalData(isAuto) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const personalSheet = ss.getSheetByName(PERSONAL_SHEET_NAME);
  const configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!personalSheet || !configSheet) return;

  const targetSheetName = personalSheet.getRange('A3').getValue(); 
  const targetStudentNo = personalSheet.getRange('B1').getValue();
  if (!targetSheetName || !targetStudentNo) return;

  const studentList = getStudentList(configSheet);
  const studentInfo = studentList.find(s => s.no == targetStudentNo);
  if (studentInfo) {
    personalSheet.getRange("E1").setValue(studentInfo.name);
    personalSheet.getRange("I1").setValue(studentInfo.email);
  } else {
    if (!isAuto) Browser.msgBox("警告", "番号不明", Browser.Buttons.OK);
    personalSheet.getRange("E1").clearContent();
    personalSheet.getRange("I1").clearContent();
  }

  const targetSheet = ss.getSheetByName(targetSheetName);
  if (!targetSheet) {
    if (!isAuto) Browser.msgBox("エラー", "シート不明", Browser.Buttons.OK);
    return;
  }
  const lastRow = targetSheet.getLastRow();
  const studentIds = targetSheet.getRange(3, 1, lastRow - 2, 1).getValues();
  let targetRow = -1;
  for (let i = 0; i < studentIds.length; i++) {
    if (studentIds[i][0] == targetStudentNo) {
      targetRow = i + 3;
      break;
    }
  }
  if (targetRow === -1) {
    personalSheet.getRange(3, 9, 31, 2).clearContent();
    if (!isAuto) ss.toast("データなし", "通知");
    return;
  }
  const daysDataRange = targetSheet.getRange(targetRow, 3, 1, 31);
  const notes = daysDataRange.getNotes()[0];
  const outputData = [];
  for (let i = 0; i < 31; i++) {
    const note = notes[i];
    let diary = "";
    let aiComment = "";
    if (note) {
      const diaryMatch = note.match(/\[日記\]\s*([\s\S]*?)\s*(\[AIコメント\]|$)/);
      if (diaryMatch) diary = diaryMatch[1].trim();
      const aiMatch = note.match(/\[AIコメント\]\s*([\s\S]*)/);
      if (aiMatch) aiComment = aiMatch[1].trim();
    }
    outputData.push([diary, aiComment]);
  }
  personalSheet.getRange(3, 9, 31, 2).setValues(outputData);
  if (isAuto) ss.toast("更新完了", "完了");
  else Browser.msgBox("完了", "読み込みました。", Browser.Buttons.OK);
}

// ==========================================
// ■ 機能3：クロス分析
// ==========================================
function runCrossAnalysis() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  const surveySheet = ss.getSheetByName(SURVEY_SHEET_NAME);
  const date = new Date();
  const year = date.getFullYear();
  const month = date.getMonth() + 1;
  const monthlySheetName = `${year}年${month}月`;
  const monthlySheet = ss.getSheetByName(monthlySheetName);

  if (!surveySheet || !monthlySheet) {
    Browser.msgBox("エラー", "シート不足", Browser.Buttons.OK);
    return;
  }
  const studentList = getStudentList(configSheet);
  const lastRowSurvey = surveySheet.getLastRow();
  const surveyData = (lastRowSurvey >= 5) ? surveySheet.getRange(5, 1, lastRowSurvey - 4, 20).getValues() : [];
  const surveyMap = {}; 
  const IDX_Q1 = 3; const IDX_Q4 = 6; const IDX_Q5 = 8; const IDX_Q6 = 10; const IDX_Q7 = 12; const IDX_MEMO_MAIN = 18; const IDX_MEMO_SUB = 16;
  surveyData.forEach(r => {
    const no = r[0]; 
    const answerText = String(r[IDX_Q1] || ""); 
    const valQ4 = String(r[IDX_Q4] || "").trim(); const valQ5 = String(r[IDX_Q5] || "").trim(); 
    const valQ6 = String(r[IDX_Q6] || "").trim(); const valQ7 = String(r[IDX_Q7] || "").trim(); 
    let memoText = String(r[IDX_MEMO_MAIN] || ""); 
    if (memoText === "") memoText = String(r[IDX_MEMO_SUB] || ""); 
    let score = 0;
    if (answerText.includes("とても楽しい") || answerText === "5") score = 5;
    else if (answerText.includes("まあまあ") || answerText === "4") score = 4;
    else if (answerText.includes("あまり") || answerText === "2") score = 2;
    else if (answerText.includes("楽しくない") || answerText === "1") score = 1;
    else if (answerText !== "") score = 3; 
    const risks = [];
    if (valQ4.includes("ある")) risks.push("いじめ被害");
    if (valQ5.includes("ある")) risks.push("金品被害");
    if (valQ6.includes("いる")) risks.push("周囲にいじめ");
    if (valQ7.includes("見た") || valQ7.includes("聞いた")) risks.push("問題行動目撃");
    if (no) surveyMap[no] = { score: score, text: answerText, risks: risks, memo: memoText };
  });

  const lastRowMonthly = monthlySheet.getLastRow();
  const daysData = monthlySheet.getRange(3, 3, lastRowMonthly - 2, 31).getValues();
  const monthlyMap = {}; 
  daysData.forEach((row, index) => {
    let sum = 0; let count = 0;
    row.forEach(cell => { if (typeof cell === 'number') { sum += cell; count++; } });
    monthlyMap[index] = count > 0 ? (sum / count) : 0;
  });

  let resultSheet = ss.getSheetByName('【分析】クロス集計');
  if (resultSheet) ss.deleteSheet(resultSheet);
  resultSheet = ss.insertSheet('【分析】クロス集計');
  const headers = ["No", "氏名", "⚠️リスク検知", "①調査回答", "②日記平均", "判定結果", "調査の記述(悩み等)", "アドバイス"];
  resultSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setBackground("#cc0000").setFontColor("white").setFontWeight("bold");
  const outputRows = [];
  studentList.forEach((s, index) => {
    const sData = surveyMap[s.no] || { score: 0, text: "未回答", risks: [], memo: "" };
    const pScore = sData.score; const dScore = monthlyMap[index] || 0; 
    let type = "-"; let advice = "-"; let color = null; 
    let riskLabel = sData.risks.join(", ");
    const hasSevereRisk = sData.risks.length > 0;
    if (hasSevereRisk) { type = "🚨 緊急SOS"; advice = `【最優先】「${riskLabel}」の回答あり。`; color = "#ea4335"; } 
    else if (pScore === 0 || dScore === 0) { type = "データ不足"; } 
    else {
      const isHighP = pScore >= 3; const isHighD = dScore >= 3;
      if (isHighP && isHighD) { type = "★ 安定・充実"; advice = "良好です。"; color = "#e6f4ea"; } 
      else if (!isHighP && !isHighD) { type = "⚠️ 要支援継続"; advice = "SOSが出ています。"; color = "#fce8b2"; } 
      else if (isHighP && !isHighD) { type = "📉 急激な悪化"; advice = "日記が悪化中。"; color = "#fff2cc"; } 
      else if (!isHighP && isHighD) { type = "🎭 仮面タイプ"; advice = "調査に不満あり。"; color = "#e8f0fe"; }
    }
    let cleanMemo = sData.memo.replace(/\n+/g, " ").trim().substring(0, 80);
    outputRows.push([s.no, s.name, riskLabel, sData.text, dScore.toFixed(1), type, cleanMemo, advice, color]);
  });
  if (outputRows.length > 0) {
    const range = resultSheet.getRange(2, 1, outputRows.length, headers.length);
    range.setValues(outputRows.map(r => r.slice(0, 8)));
    const colors = outputRows.map(r => Array(headers.length).fill(r[8] || "white"));
    range.setBackgrounds(colors);
    resultSheet.getRange(2, 1, outputRows.length, headers.length).setVerticalAlignment("middle").setWrap(true);
  }
  Browser.msgBox("分析完了", "完了しました。", Browser.Buttons.OK);
}

// ==========================================
// ■ 機能4：単体PDF作成（#REF!対策版）
// ==========================================
function createPdf() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PERSONAL_SHEET_NAME);
  if (!sheet) { Browser.msgBox("エラー", "シート不明", Browser.Buttons.OK); return; }

  const month = sheet.getRange("A3").getValue(); 
  const studentName = sheet.getRange("E1").getValue(); 
  const filename = `${month}_${studentName}_健康観察.pdf`;

  SpreadsheetApp.flush();
  Utilities.sleep(2000); // 描画待ち

  try {
    ss.toast("PDF作成中...", "処理中");
    
    // 1. 現在のシートを同じファイル内にコピー（これでリンク切れ回避！）
    const tempSheet = sheet.copyTo(ss);
    tempSheet.setName("TEMP_PDF_SHEET");

    // 2. コピーしたシートの数式を「値」に変換（#REF!回避の核心）
    const dataRange = tempSheet.getDataRange();
    dataRange.setValues(dataRange.getValues());

    // 3. 不要部分を隠す（A1:J54のみ表示）
    const maxRows = tempSheet.getMaxRows();
    const maxCols = tempSheet.getMaxColumns();
    if (maxRows > 54) tempSheet.hideRows(55, maxRows - 54);
    if (maxCols > 10) tempSheet.hideColumns(11, maxCols - 10);
    SpreadsheetApp.flush();

    // 4. PDF出力
    const blob = getSheetAsPdfBlob(ss, tempSheet, filename);
    const file = DriveApp.createFile(blob);

    // 5. 後始末（一時シート削除）
    ss.deleteSheet(tempSheet);

    showSuccessDialog(file.getUrl(), DriveApp.getRootFolder().getUrl(), "作成完了（#REF解消版）");

  } catch (e) {
    Browser.msgBox("エラー", "失敗: " + e.message, Browser.Buttons.OK);
  }
}

// ==========================================
// ■ 機能5：一括PDF作成（#REF!対策版）
// ==========================================
function createMergedPdfFromSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PERSONAL_SHEET_NAME);
  const ui = SpreadsheetApp.getUi();
  if (!sheet) { ui.alert("エラー", "シート不明", ui.ButtonSet.OK); return; }

  const startNum = sheet.getRange("M16").getValue(); 
  const endNum = sheet.getRange("M17").getValue();   
  if (typeof startNum !== 'number' || typeof endNum !== 'number') {
    ui.alert('エラー', 'M16, M17に数値を入力してください。', ui.ButtonSet.OK); return;
  }
  if (ui.alert('確認', `${startNum}番～${endNum}番を作成しますか？`, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  try {
    ss.toast("一括作成中...", "開始", -1);
    const parentFolder = DriveApp.getFileById(ss.getId()).getParents().next();
    const monthStr = sheet.getRange("A3").getValue();
    const tempSsName = `【一括】${monthStr}_No.${startNum}-${endNum}_健康観察`;
    const tempSs = SpreadsheetApp.create(tempSsName);
    const tempSsId = tempSs.getId();
    const defaultSheet = tempSs.getSheets()[0];

    for (let i = startNum; i <= endNum; i++) {
      sheet.getRange("B1").setValue(i);
      loadPersonalData(true); 
      SpreadsheetApp.flush();
      Utilities.sleep(2000); 

      const studentName = sheet.getRange("E1").getValue() || `No.${i}`;
      const values = sheet.getDataRange().getValues();

      // 直接一時ファイルへシートをコピー（ssへの中間コピー・削除のオーバーヘッドを削減）
      const copiedSheet = sheet.copyTo(tempSs);
      
      // コピー先の数式を値に置き換えて固定
      copiedSheet.getRange(1, 1, values.length, values[0].length).setValues(values);
      
      // シート名を設定（不適切な文字を除去して30文字以内に制限）
      const cleanName = `${i}.${studentName}`.replace(/[\[\]\*\?\:\\\/]/g, '').substring(0, 30);
      copiedSheet.setName(cleanName);
      
      // 印刷範囲の調整
      const maxRows = copiedSheet.getMaxRows();
      const maxCols = copiedSheet.getMaxColumns();
      if (maxRows > 54) copiedSheet.hideRows(55, maxRows - 54);
      if (maxCols > 10) copiedSheet.hideColumns(11, maxCols - 10);
    }

    sheet.getRange("B1").setValue(startNum);
    loadPersonalData(true); 
    tempSs.deleteSheet(defaultSheet);
    
    // PDF化
    const url_ext = 'export?exportFormat=pdf&format=pdf&size=A4&portrait=true&fitw=true&sheetnames=false&printtitle=false&pagenumbers=true&gridlines=false&fzr=false';
    const token = ScriptApp.getOAuthToken();
    const resp = UrlFetchApp.fetch(`https://docs.google.com/spreadsheets/d/${tempSsId}/${url_ext}`, { headers: { 'Authorization': 'Bearer ' + token } });
    const blob = resp.getBlob().setName(tempSsName + ".pdf");
    const savedFile = parentFolder.createFile(blob);
    DriveApp.getFileById(tempSsId).setTrashed(true);

    showSuccessDialog(savedFile.getUrl(), parentFolder.getUrl(), "一括作成完了");

  } catch (e) {
    ui.alert('エラー', e.message, ui.ButtonSet.OK);
  }
}

// ▼▼ ヘルパー ▼▼
function getSheetAsPdfBlob(ss, sheet, filename) {
  const url = "https://docs.google.com/spreadsheets/d/" + ss.getId() + "/export?" +
    "format=pdf&gid=" + sheet.getSheetId() + "&size=A4&portrait=true&fitw=true&gridlines=false&printtitle=false&sheetnames=false&pagenum=false";
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + token } });
  return response.getBlob().setName(filename);
}

function showSuccessDialog(fileUrl, folderUrl, title) {
  const htmlOutput = HtmlService.createHtmlOutput(
    `<div style="font-family: sans-serif; padding: 10px;"><p><b>${title}</b></p><ul><li><a href="${fileUrl}" target="_blank">📄 PDFを開く</a></li><li><a href="${folderUrl}" target="_blank">📂 保存先を開く</a></li></ul><p><button onclick="google.script.host.close()">閉じる</button></p></div>`
  ).setWidth(350).setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, title);
}
// ※getStudentMap, getStudentList, callGeminiForMental, writeToMonthlySheet, registerSheetToConfig は変更なし（省略していません、そのまま使えます）
function getStudentMap(sheet) { const lastRow = sheet.getLastRow(); const rangeHeight = lastRow - (START_ROW_STUDENT - 1); if (rangeHeight < 1) return {}; const data = sheet.getRange(START_ROW_STUDENT, 1, rangeHeight, 3).getValues(); const map = {}; data.forEach(row => { if(row[1]) map[row[1].toString().trim().toLowerCase()] = { no: row[0], name: row[2] }; }); return map; }
function getStudentList(sheet) { const lastRow = sheet.getLastRow(); const rangeHeight = lastRow - (START_ROW_STUDENT - 1); if (rangeHeight < 1) return []; const data = sheet.getRange(START_ROW_STUDENT, 1, rangeHeight, 3).getValues(); const list = []; data.forEach(row => { if(row[1]) list.push({ no: row[0], name: row[2], email: row[1].toString().trim().toLowerCase() }); }); return list; }
function callGeminiForMental(apiKey, name, selfScore, diary) { const modelName = 'gemini-2.5-flash'; const apiUrl = `https://generativelanguage.googleapis.com/v1/models/${modelName}:generateContent?key=${apiKey}`; const prompt = `あなたは学校カウンセラー支援AIです。JSON形式のみで出力。生徒:${name},自己評価:${selfScore},日記:"${diary}"。分析ルール:1(SOS)-5(元気)。出力JSON:{ "comment": "30字以内要約", "sosLevel": 数値(1-5) }`; const payload = { "contents": [{ "parts": [{ "text": prompt }] }], "generationConfig": { "response_mime_type": "application/json" } }; const options = { "method": "post", "contentType": "application/json", "payload": JSON.stringify(payload), "muteHttpExceptions": true }; const response = UrlFetchApp.fetch(apiUrl, options); if (response.getResponseCode() !== 200) throw new Error(response.getContentText()); const json = JSON.parse(response.getContentText()); return JSON.parse(json.candidates[0].content.parts[0].text); }
function writeToMonthlySheet(ss, studentList, date, student, score, diary, aiResult) { const year = date.getFullYear(); const month = date.getMonth() + 1; const sheetName = `${year}年${month}月`; const day = date.getDate(); let sheet = ss.getSheetByName(sheetName); if (!sheet) { sheet = ss.insertSheet(sheetName); sheet.getRange(1, 1).setValue("No"); sheet.getRange(1, 2).setValue("氏名"); const listValues = studentList.map(s => [s.no, s.name]); if (listValues.length > 0) sheet.getRange(3, 1, listValues.length, 2).setValues(listValues); const lastRow = listValues.length + 2; sheet.getRange(1, 1, lastRow, 33).setBackground(null); const daysOfWeek = ["日", "月", "火", "水", "木", "金", "土"]; for (let d = 1; d <= 31; d++) { const colIndex = d + 2; const currentDate = new Date(year, month - 1, d); if (currentDate.getMonth() !== month - 1) break; const dayOfWeek = currentDate.getDay(); const dayChar = daysOfWeek[dayOfWeek]; sheet.getRange(1, colIndex).setValue(d); sheet.getRange(2, colIndex).setValue(dayChar); sheet.setColumnWidth(colIndex, 25); if (dayOfWeek === 0 || dayOfWeek === 6) { sheet.getRange(1, colIndex, lastRow, 1).setBackground("#e0e0e0"); } } sheet.getRange(1, 1, lastRow, 33).setBorder(true, true, true, true, true, true); sheet.setFrozenRows(2); sheet.setFrozenColumns(2); sheet.setColumnWidth(1, 40); sheet.setColumnWidth(2, 150); const range = sheet.getRange(3, 3, studentList.length, 31); const rules = sheet.getConditionalFormatRules(); rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberEqualTo(1).setBackground("#ea4335").setFontColor("white").setRanges([range]).build()); rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberEqualTo(2).setBackground("#fa7b17").setRanges([range]).build()); rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberEqualTo(3).setBackground("#fbbc04").setRanges([range]).build()); rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(4, 5).setBackground("#e8f0fe").setRanges([range]).build()); sheet.setConditionalFormatRules(rules); registerSheetToConfig(ss, sheet); } const studentIndex = studentList.findIndex(s => s.no == student.no && s.name == student.name); if (studentIndex === -1) return; const cell = sheet.getRange(studentIndex + 3, day + 2); cell.setValue(aiResult.sosLevel); cell.setNote(`【気分: ${score} / 評価: ${aiResult.sosLevel}】\n\n[日記]\n${diary}\n\n[AIコメント]\n${aiResult.comment}`); }
function registerSheetToConfig(ss, newSheet) { const configSheet = ss.getSheetByName(CONFIG_SHEET_NAME); if (!configSheet) return; const sheetName = newSheet.getName(); const sheetId = newSheet.getSheetId(); const fullUrl = `${ss.getUrl()}#gid=${sheetId}`; const startRow = 3; const maxCheckRows = 200; const range = configSheet.getRange(startRow, 5, maxCheckRows, 1); const values = range.getValues(); let targetRow = -1; for (let i = 0; i < values.length; i++) { const val = String(values[i][0]); if (val === sheetName || val === "'" + sheetName) return; if (val === "" && targetRow === -1) { targetRow = startRow + i; break; } } if (targetRow !== -1) { configSheet.getRange(targetRow, 5).setValue("'" + sheetName); configSheet.getRange(targetRow, 6).setFormula(`=HYPERLINK("${fullUrl}", "リンク")`); } }