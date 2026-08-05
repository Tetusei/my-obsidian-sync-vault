/**
 * 【ファイル3】Kondate.gs (v3.1.5)
 * 献立表の解析：A列日付・B列献立（2段改行・分離問題解決版）
 */
function processKondate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.CONFIG);
  const apiKey = configSheet.getRange('B2').getValue();
  const inputFolderId = extractId(configSheet.getRange('B3').getValue());
  const year = getCurrentSchoolYear(); 

  writeLog("🚀 献立解析プロセス開始", "info");
  
  if (!apiKey || !inputFolderId) {
    writeLog("⚠️ 設定エラー：APIキーまたはフォルダIDが不足しています", "error");
    return;
  }

  const file = findFileByKeyword(inputFolderId, "献立");
  if (!file) {
    writeLog("⚠️ 「献立」PDFが見つかりません", "error");
    return;
  }

  try {
    writeLog(`📄 ${file.getName()} を読み込んでいます...`, "info");
    
    // 🌟 修正ポイント：2段改行を1つの料理名として繋げるようプロンプトを強化
    const prompt = `
あなたは学校給食のデータ入力専門家です。
PDFから「日付」と「献立名」を抽出してください。

【重要：献立名の抽出ルール】
1. セル内で2段以上に改行されている献立名（例：「スタミナ」と「ライス」や、「コーン」と「ミルクスープ」）は、絶対に分離せず、1つの繋がった料理名（「スタミナライス」「コーンミルクスープ」など）として抽出してください。
2. 同一日に複数の献立がある場合は、スペース等で区切って1つの文字列にまとめてください。

【日付の扱い】
紙面の年号（令和8年など）を絶対優先。設定値(${year})は無視して西暦変換してください。

【出力形式】
YYYY/MM/DD|献立名
`;
    
    const text = callGeminiWithRetry(apiKey, file.getBlob(), prompt);
    const rows = parseGeminiOutput(text, 2);
   
    if (rows.length === 0) {
      writeLog("❌ データの抽出に失敗しました", "error");
      throw new Error("データ抽出失敗");
    }

    let sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MASTER);
    const TARGET_START_ROW = 3;
    
    const lastRow = sheet.getLastRow();
    if (lastRow >= TARGET_START_ROW) {
      sheet.getRange(TARGET_START_ROW, 1, lastRow - TARGET_START_ROW + 1, 4).clearContent();
    }

    const finalRows = rows.map(r => [
      r[0],      // A列：日付
      r[1],      // B列：献立名
      "",        // C列：空（ここに材料解析の結果が入る）
      ""         // D列：空
    ]);

    sheet.getRange(TARGET_START_ROW, 1, finalRows.length, 4).setValues(finalRows);
    sheet.getRange(TARGET_START_ROW, 1, finalRows.length, 1).setNumberFormat("yyyy/mm/dd");
    const dataRange = sheet.getRange(TARGET_START_ROW, 1, finalRows.length, 4);
    dataRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP).setVerticalAlignment("top");
    
    const existing = sheet.getRange(1, 1, sheet.getMaxRows(), 4).getBandings();
    existing.forEach(b => b.remove());
    sheet.getRange(TARGET_START_ROW - 1, 1, finalRows.length + 1, 4).applyRowBanding(SpreadsheetApp.BandingTheme.GREEN);
   
    writeLog(`✅ 献立完了: ${finalRows.length}件の日付と献立名を登録しました`, "success");
    
  } catch (e) {
    writeLog(`❌ 献立エラー: ${e.message}`, "error");
  }
}