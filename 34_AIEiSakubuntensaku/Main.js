/**
 * ファイル名: Main.gs
 * 役割: AI添削の実行・API通信・進捗更新
 * バージョン: v4.0.0
 */

const Main = {
  updateProgress: function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(Config.MAIN_SHEET);
    if (!sheet) return;
    
    // 分母をA列（生徒番号）でカウントする最強の式
const formula = '="📊 現在の提出状況: " & COUNTA(C3:C) & " / " & COUNTA(A3:A) & " 人 (" & TEXT(IF(COUNTA(A3:A)=0, 0, COUNTA(C3:C)/COUNTA(A3:A)), "0%") & ")"';
    sheet.getRange('D1').setFormula(formula);
  },

  processUnprocessedRows: function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(Config.MAIN_SHEET);
    if (!sheet) {
      SpreadsheetApp.getUi().alert(`⚠️ 「${Config.MAIN_SHEET}」シートが見つかりません。`);
      return;
    }

    let apiKeys = [], customPrompt = "", actionMode = ""; 
    try {
      apiKeys = getApiKeys(); customPrompt = getCustomPrompt(); actionMode = getActionMode(); 
    } catch(e) { SpreadsheetApp.getUi().alert('⚠️ 設定エラー\n' + e.message); return; }

    const questionText = sheet.getRange(Config.QUESTION_CELL).getValue() || "";

    const lastRow = sheet.getLastRow();
    if (lastRow < Config.DATA_START_ROW) return;
    const dataRange = sheet.getRange(Config.DATA_START_ROW, 1, lastRow - Config.DATA_START_ROW + 1, Config.COL_SCORE); 
    const data = dataRange.getValues();
    
    let successCount = 0, errorCount = 0, keyIndex = 0; 

    for (let i = 0; i < data.length; i++) {
      const rowNum = i + Config.DATA_START_ROW;
      const studentText = data[i][Config.COL_STUDENT - 1]; 
      const currentFix = data[i][Config.COL_FIX - 1] || ''; 
      
      if (!studentText) continue;
      
      const needsProcess = (!currentFix || currentFix.toString().includes('エラー') || currentFix.toString().includes('生成中'));

      if (needsProcess) {
        ss.toast(`現在 ${rowNum} 行目をAIが処理しています...`, '🤖 処理中', -1);
        sheet.getRange(rowNum, Config.COL_FIX, 1, 4).setValues([['（生成中...）', '（生成中...）', '（生成中...）', '（生成中...）']]);
        SpreadsheetApp.flush(); 

        const result = Main.callGeminiUnified_(apiKeys, keyIndex, Main.getPromptUnified(studentText, customPrompt, actionMode, questionText));
        keyIndex++;

        if (result.error) {
          sheet.getRange(rowNum, Config.COL_FIX, 1, 4).setValues([['エラー', 'エラー', 'エラー', 'エラー']]);
          errorCount++; 
        } else {
          let finalScore = result.score || '-';
          if (finalScore !== '-' && !finalScore.toString().includes('点') && finalScore.toString().includes('/')) {
            finalScore = finalScore + '点';
          }
          if (finalScore !== '-') finalScore = "'" + finalScore;

          sheet.getRange(rowNum, Config.COL_FIX, 1, 4).setValues([[
            result.fixed_text || '-', 
            result.explanation || '-', 
            result.japanese_translation || '-',
            finalScore
          ]]);
          successCount++; 
        }
        Utilities.sleep(1000); 
      }
    }
    
    if (successCount > 0) {
      ss.toast(`${successCount}件の処理が成功しました。`, '✨ 完了', 5);
    } else if (errorCount > 0) {
      ss.toast(`${errorCount}件のエラーが発生しました。時間を置いて再度実行してください。`, '⚠️ エラー', 5);
    } else {
      ss.toast('処理が必要なデータはありませんでした。', '✅ 完了', 3);
    }
  },

  getPromptUnified: function(text, customInstructions, actionMode, questionText) {
      let instructionText = customInstructions ? customInstructions.trim() : "特になし。中学生が習う範囲の語彙や文法を使って自然な英語に直してください。";
      let qTextForPrompt = questionText ? questionText : "（特に指定されていません。生徒の英文から意図を推測してください）";

      if (actionMode && actionMode.indexOf('ヒント') !== -1) {
        return `
あなたは親切な中学校の英語教師です。生徒が自分で間違いに気づけるようヒントだけを出します。
【特別指示】${instructionText}
【生徒が英語にしたかった元の日本語（問題文）】: ${qTextForPrompt}
【ルール】
1. fixed_text: 「（自分で直してみよう！）」を出力。完璧な場合は「（お見事！修正の必要はありません✨）」
2. explanation: 元の日本語の意図を踏まえ、どこが間違っているかのヒントのみ。完璧な場合は褒め言葉。
3. japanese_translation: 「（自分の文をもう一度読んでみよう）」を出力。完璧な場合は和訳。
4. score: ヒントモードなので「- (採点なし)」と出力。
【生徒の英文】: ${text}
必ず以下のJSONで出力: {"fixed_text":"", "explanation":"", "japanese_translation":"", "score":""}`;
      } else {
        return `
あなたは親切な中学校の英語教師です。
【特別指示】${instructionText}
【生徒が英語にしたかった元の日本語（問題文）】: ${qTextForPrompt}
【ルール】
1. fixed_text: 元の日本語の意図を汲み取りつつ、指示に従い、中学生にふさわしい英語に添削。
2. explanation: なぜそう直したのか3行以内で解説。元の日本語のニュアンスにどう合わせるかの解説も含める。
3. japanese_translation: 添削後の自然な和訳。
4. score: 生徒の元の英文を10点満点で採点し、「8/10点」のように出力。文法、語彙、問題文への合致度を総合的に評価。完璧なら「10/10点」。
【生徒の英文】: ${text}
必ず以下のJSONで出力: {"fixed_text":"", "explanation":"", "japanese_translation":"", "score":""}`;
      }
  },

  callGeminiUnified_: function(apiKeys, startIndex, prompt) {
    const payload = { "contents": [{ "parts": [{ "text": prompt }] }], "generationConfig": { "responseMimeType": "application/json" } };
    const maxRetries = Math.max(3, apiKeys.length * 2);
    for (let i = 0; i < maxRetries; i++) {
      const currentKey = apiKeys[(startIndex + i) % apiKeys.length];
      const url = `https://generativelanguage.googleapis.com/v1beta/models/${Config.MODEL_NAME}:generateContent?key=${currentKey}`;
      const options = { "method": "post", "contentType": "application/json", "payload": JSON.stringify(payload), "muteHttpExceptions": true };
      try {
        const response = UrlFetchApp.fetch(url, options);
        if (response.getResponseCode() === 429) { Utilities.sleep(2000); continue; }
        const json = JSON.parse(response.getContentText());
        if (json.candidates) {
          let text = json.candidates[0].content.parts[0].text.trim();
          if (text.startsWith('```')) text = text.replace(/^```(json)?\n?/, '').replace(/\n?```$/, '');
          try { return JSON.parse(text); } catch (e) { return { error: "解析エラー" }; }
        } else if (json.error) {
          if (json.error.message.includes("quota") || json.error.message.includes("429")) { Utilities.sleep(2000); continue; }
          return { error: json.error.message };
        } 
      } catch (e) { if (e.toString().includes("429") || e.toString().includes("Quota")) { Utilities.sleep(2000); continue; } return { error: e.toString() }; }
    }
    return { error: "回線混雑中" };
  }
};