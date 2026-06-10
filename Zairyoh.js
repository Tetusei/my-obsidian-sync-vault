/**
 * 【ファイル4】Zairyoh.gs (v3.5.2)
 * 役割：フォルダ内の「材料」を含むすべてのPDFを順番に解析し、料理名(品名)付きでマスタに安全に転記する
 */
function processZairyoh() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.CONFIG);
  const apiKey = configSheet.getRange('B2').getValue();
  const inputFolderId = extractId(configSheet.getRange('B3').getValue());

  if (typeof writeLog === 'function') writeLog("🚀 材料解析プロセスを開始しました（複数ファイル一括処理）", "info");

  const files = findAllFilesByKeyword(inputFolderId, "材料");
  if (!files || files.length === 0) {
    if (typeof writeLog === 'function') writeLog("⚠️ エラー：指定フォルダ内に「材料」という名前のPDFが見つかりません。", "error");
    return;
  }
  
  if (typeof writeLog === 'function') writeLog(`📂 フォルダ内に ${files.length} 件の「材料」ファイルを確認しました。順次解析します。`, "info");

  try {
    const masterSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MASTER);
    const lastRow = masterSheet.getLastRow();
    if (lastRow < CONFIG.START_ROW) {
      if (typeof writeLog === 'function') writeLog("❌ エラー：献立マスタに日付データがありません。", "error");
      return;
    }

    let dictionaryText = "";
    const dictSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.DICT);
    if (dictSheet && dictSheet.getLastRow() >= CONFIG.START_ROW) {
      const dictData = dictSheet.getRange(CONFIG.START_ROW, 1, dictSheet.getLastRow() - CONFIG.START_ROW + 1, 3).getValues();
      dictionaryText = dictData
        .filter(r => r[0] && r[1])
        .map(r => `・${r[0]}：${r[1]}${r[2] ? "（除外："+r[2]+"）" : ""}`)
        .join("\n");
    }

    const baseSheet = ss.getSheetByName("基礎データ");
    if (!baseSheet) {
      if (typeof writeLog === 'function') writeLog("❌ エラー：「基礎データ」シートが見つかりません。", "error");
      return;
    }
    let userPrompt = baseSheet.getRange("B12").getValue();
    if (!userPrompt) {
      if (typeof writeLog === 'function') writeLog("❌ エラー：基礎データシートのB12セルが空です。指示文を入力してください。", "error");
      return;
    }

    // 🌟 修正ポイント：AIが確実に【品名】材料 の形式で出力するようにシステム側から命令を強制挿入
    const finalPrompt = `${userPrompt}

【絶対厳守する抽出・出力ルール】
1. PDFから「日付」「品名（料理名）」「使用されている原材料」を確実に抽出してください。
2. 出力は、1日分のデータを「1行」にまとめてください。
3. 2列目（材料列）は、必ず「【料理名1】材料A、材料B 【料理名2】材料C、材料D」の形式で出力してください。抽出した品名（料理名）を【 】で囲み、その直後に原材料を続けてください。

--- 判定辞書（A列が分類名、B列が単語） ---
${dictionaryText || "なし"}`;

    const masterRange = masterSheet.getRange(CONFIG.START_ROW, 1, lastRow - CONFIG.START_ROW + 1, 4);
    const masterData = masterRange.getValues();
    let totalMatchCount = 0;

    for (let f = 0; f < files.length; f++) {
      const file = files[f];
      if (typeof writeLog === 'function') writeLog(`🧠 AI処理中 (${f+1}/${files.length})：『${file.getName()}』を解析しています...`, "info");
      
      let text = "";
      try {
        text = callGeminiWithRetry(apiKey, file.getBlob(), finalPrompt);
      } catch (err) {
        if (typeof writeLog === 'function') writeLog(`❌ 『${file.getName()}』の解析中にエラー: ${err.message}`, "error");
        continue;
      }

      const rows = parseGeminiOutput(text, 2); 
      
      if (rows.length === 0) {
        if (typeof writeLog === 'function') writeLog(`⚠️ 警告：『${file.getName()}』からはデータが抽出できませんでした。`, "warn");
        continue;
      }

      let fileMatchCount = 0;

      rows.forEach(row => {
        try {
          let aiM = -1, aiD = -1;
          const rawDateStr = String(row[0]).trim();

          const tempDate = new Date(rawDateStr);
          if (!isNaN(tempDate.getTime())) {
            aiM = tempDate.getMonth() + 1;
            aiD = tempDate.getDate();
          } else {
            const match = rawDateStr.match(/(?:^|\D)(\d{1,2})[/\-月](\d{1,2})(?:日|$|\D)/);
            if (match) {
              aiM = parseInt(match[1], 10);
              aiD = parseInt(match[2], 10);
            }
          }

          if (aiM === -1 || aiD === -1) return;

          for (let i = 0; i < masterData.length; i++) {
            const mDateVal = masterData[i][0];
            if (!mDateVal) continue;

            const mDateObj = (mDateVal instanceof Date) ? mDateVal : new Date(mDateVal);
            if (isNaN(mDateObj.getTime())) continue;

            if (aiM === (mDateObj.getMonth() + 1) && aiD === mDateObj.getDate()) {
              
              let isUpdated = false;
              // 🌟 鉄壁の安全装置：既存のデータがある場合は「追記」して消さないようにする
              if (row[1] && String(row[1]).trim() !== "") {
                const currentMat = String(masterData[i][2] || "").trim();
                if (currentMat && !currentMat.includes(row[1])) {
                   masterData[i][2] = currentMat + " " + row[1]; 
                } else {
                   masterData[i][2] = row[1];
                }
                isUpdated = true;
              }
              if (row[2] && String(row[2]).trim() !== "") {
                const currentJudge = String(masterData[i][3] || "").trim();
                if (currentJudge && !currentJudge.includes(row[2])) {
                   masterData[i][3] = currentJudge + "\n" + row[2];
                } else {
                   masterData[i][3] = row[2];
                }
                isUpdated = true;
              }
              
              if (isUpdated) {
                 fileMatchCount++;
                 totalMatchCount++;
              }
              break;
            }
          }
        } catch(e) { console.error("Match Error: " + e.message); }
      });

      if (typeof writeLog === 'function') writeLog(` ↳ 『${file.getName()}』から ${fileMatchCount} 件のデータを準備しました。`, "success");

      if (f < files.length - 1) {
        Utilities.sleep(5000); 
      }
    } 

    if (totalMatchCount > 0) {
      masterRange.setValues(masterData);
      if (typeof writeLog === 'function') writeLog(`✅ 全ファイルの解析完了！ 合計 ${totalMatchCount} 件のデータを更新しました。`, "success");
      
      if (typeof applySpecificAiColor === 'function') applySpecificAiColor(); 
    } else {
      if (typeof writeLog === 'function') writeLog("⚠️ 有効データなし：新しいデータがなかったため、元のデータを保護しました。", "warn");
    }

  } catch (e) {
    if (typeof writeLog === 'function') writeLog(`❌ システムエラー：${e.message}`, "error");
  }
}