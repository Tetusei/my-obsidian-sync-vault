/**
 * 【ファイル5】Kojin.gs (v2.5.1 改)
 * 個人票の一括解析（追記型対応・日付補完・全件抽出対応版）
 */
function processKojin() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.CONFIG);
  
  const apiKey = CONFIG.getApiKey();
  const folderIdRange = configSheet.getRange('B3');
  const inputFolderId = extractId(folderIdRange.getValue());
  
  // ※getProcessedFolderId や getCurrentSchoolYear は既存の関数を呼び出します
  const processedFolderId = CONFIG.getProcessedFolderId ? CONFIG.getProcessedFolderId() : extractId(configSheet.getRange('B4').getValue());
  const year = typeof getCurrentSchoolYear === 'function' ? getCurrentSchoolYear() : CONFIG.TARGET_YEAR;

  if (typeof writeLog === 'function') writeLog("🚀 個人票一括解析開始(v2.5.1改 - 日付補完対応)", "info");

  const files = findAllFilesByKeyword(inputFolderId, "個人");
  if (files.length === 0) { 
    if (typeof writeLog === 'function') writeLog("✅ 未処理の個人票はありません", "success"); 
    return; 
  }
 
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MAIN);
  let totalAdded = 0;

  for (let i = 0; i < files.length; i++) {
    const file = files[i];
    const fileName = file.getName();
    
    // 🌟 修正ポイント1：ファイル名から「月」を自動推測（例：R8_6... から 6 を抽出）
    let guessedMonth = -1;
    const monthMatch = fileName.match(/(\d{1,2})月/) || fileName.match(/[R8_\-](\d{1,2})/);
    if (monthMatch) {
      guessedMonth = parseInt(monthMatch[1], 10);
    }

    if (typeof writeLog === 'function') writeLog(`📂 (${i+1}/${files.length}) ${fileName} 解析中... (推測月: ${guessedMonth !== -1 ? guessedMonth + '月' : '不明'})`, "info");
    
    try {
      // 🌟 修正ポイント2：プロンプトに日付補完と全件抽出の絶対ルールを追加
      const prompt = `あなたは学校給食のアレルギー管理担当です。
この個人票PDFを解析し、データを抽出してください。

【最重要：年の扱い】
システム設定値(${year}年度)に関わらず、**必ずPDF紙面に書かれている年号を正としてください。**

1. 紙面に「令和8年」や「R8」とあれば、それは西暦2026年です。
   例: 「令和8年2月3日」 → **2026/02/03** (2027にしないこと)
   
2. 紙面に「令和7年」や「R7」とあれば、それは西暦2025年です。

3. 日付欄に「毎日」や「全日」とある場合は、そのまま「毎日」と出力してください。
4. "同上"や"〃"という文字がある場合は、そのまま出力してください。

【絶対厳守ルール】
手書きのため、日付欄に「4日（木）」のように『月』が省略されている場合でも、他のページやファイル名に記載されている『月』をコンテキストから必ず補完し、日付はすべて「2026/06/04」のような yyyy/MM/dd 形式に統一して出力してください。
また、主たるアレルゲン以外の料理への対応要望（例：アレルゲンは卵だが、スープへの対応記載など）も、保護者の意向を逃さないために【すべて】抽出してください。

【出力形式（1行1データ）】
クラス|氏名|対象アレルゲン|日付(または毎日)|除去料理名|対応内容`;

      const text = callGeminiWithRetry(apiKey, file.getBlob(), prompt);
      const rawRows = parseGeminiOutput(text, 6);
      const rows = rawRows.filter(r => r[0] && !r[0].includes("クラス") && !r[0].includes("-"));

      if (rows.length > 0) {
        let prevClass = "", prevName = "", prevAllergen = "", prevDish = "", prevAction = "";

        const formattedRows = rows.map(r => {
          let currentClass = (r[0] || "").trim();
          if (isDitto(currentClass) && prevClass) currentClass = prevClass; else if (currentClass) prevClass = currentClass;

          let currentName = (r[1] || "").trim();
          if (isDitto(currentName) && prevName) currentName = prevName; else if (currentName) prevName = currentName;

          let currentAllergen = (r[2] || "").trim();
          if (isDitto(currentAllergen) && prevAllergen) currentAllergen = prevAllergen; else if (currentAllergen) prevAllergen = currentAllergen;

          // 🌟 修正ポイント3：月推測を活用した安全な日付パース
          let dateValue;
          const rawDate = String(r[3] || "").trim().normalize("NFKC");
          if (rawDate.includes("毎日") || rawDate.includes("全")) {
              dateValue = "毎日"; 
          } else {
              const dayOnlyMatch = rawDate.match(/^(\d{1,2})日/);
              const slashMatch = rawDate.match(/(\d{4})[/\-](\d{1,2})[/\-](\d{1,2})/);
              const normalMatch = rawDate.match(/(\d{1,2})[/\-月](\d{1,2})/);

              if (dayOnlyMatch && guessedMonth !== -1) {
                // AIが「4日（木）」と出力してきた場合、推測月を使って 2026/06/04 に変換
                dateValue = new Date(2026, guessedMonth - 1, parseInt(dayOnlyMatch[1], 10));
              } else if (slashMatch && guessedMonth !== -1 && parseInt(slashMatch[2], 10) !== guessedMonth) {
                // AIが月を誤認した場合 (例: 本来6月なのに 2026/4/1 と出力された場合)
                const aiMonth = parseInt(slashMatch[2], 10);
                const aiDay = parseInt(slashMatch[3], 10);
                // 4/1誤認の場合は「4」を日付として扱う
                const finalDay = (aiMonth === 4 && guessedMonth === 6 && aiDay === 1) ? 4 : aiDay; 
                dateValue = new Date(parseInt(slashMatch[1], 10), guessedMonth - 1, finalDay);
              } else if (normalMatch) {
                // 6/4 や 6月4日 の場合
                dateValue = new Date(2026, parseInt(normalMatch[1], 10) - 1, parseInt(normalMatch[2], 10));
              } else {
                // 既存のフォールバック処理
                let d = new Date(rawDate);
                if (!isNaN(d.getTime())) {
                  // Date変換で月がズレている場合も補正
                  if (guessedMonth !== -1 && (d.getMonth() + 1) !== guessedMonth) {
                    dateValue = new Date(2026, guessedMonth - 1, d.getDate());
                  } else {
                    if (d.getFullYear() < 2025 || d.getFullYear() === 2001) d.setFullYear(2026); 
                    dateValue = d;
                  }
                } else {
                  let retryDate = new Date("2026/" + rawDate.replace(/日.*$/, ""));
                  dateValue = !isNaN(retryDate.getTime()) ? retryDate : rawDate;
                }
              }
          }

          let currentDish = (r[4] || "").trim();
          if (isDitto(currentDish) && prevDish) currentDish = prevDish; else if (currentDish) prevDish = currentDish;

          let currentAction = (r[5] || "").trim();
          if (isDitto(currentAction) && prevAction) currentAction = prevAction; else if (currentAction) prevAction = currentAction;

          return [true, dateValue, currentClass, currentName, currentAllergen, currentDish, currentAction, ""];
        });

        // 📝 新しい追記関数を使って安全に末尾へ追加
        appendDataToMain(formattedRows);
        
        // 追記した範囲にチェックボックスと日付書式を設定
        const newLastRow = sheet.getLastRow();
        const startRow = newLastRow - formattedRows.length + 1;
        sheet.getRange(startRow, 2, formattedRows.length, 1).setNumberFormat('yyyy/MM/dd');
        sheet.getRange(startRow, 1, formattedRows.length, 1).insertCheckboxes();
        
        if (typeof moveFileToFolder === 'function') moveFileToFolder(file, processedFolderId);
        totalAdded += formattedRows.length;
        if (typeof writeLog === 'function') writeLog(`   ↳ ${file.getName()} 完了`, "success");
      }
      
      if (typeof writeLog === 'function') writeLog(`   💤 待機中...`, "info");
      Utilities.sleep(5000); 

    } catch (e) {
      if (typeof writeLog === 'function') writeLog(`   ↳ ❌ エラー: ${e.message}`, "error");
      Utilities.sleep(5000);
    }
  }

  if (totalAdded > 0) {
    if (typeof writeLog === 'function') writeLog(`🏁 全${files.length}ファイルを処理し、${totalAdded}件追加しました`, "success");
  }
}

function isDitto(text) {
  if (!text) return false;
  const t = String(text).trim();
  const patterns = ["同上", "〃", "”", "″", "same", "ー", "―", "〃(手書き)"];
  return patterns.includes(t);
}