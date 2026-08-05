/**
 * Utils.gs
 * Description: AI通信、セキュリティチェックなどの裏方機能
 * Version: v2.5.0 (リトライ処理・安全フィルター自動設定を追加)
 */

function callGeminiAPI(prompt, baseSheet, isJson = false) {
  const apiKeys = [];
  const scriptProperties = PropertiesService.getScriptProperties();

  CONFIG.CELL_API_KEYS.forEach((cell, index) => {
    const cellValue = baseSheet.getRange(cell).getValue().toString().trim();
    const propKey = `GEMINI_API_KEY_${index + 1}`;

    if (cellValue === "設定済み") {
      const savedKey = scriptProperties.getProperty(propKey);
      if (savedKey) {
        apiKeys.push(savedKey);
      }
    } else if (cellValue !== "") {
      // セルに直接生のAPIキーが残っている場合は、プロパティに格納してセルを「設定済み」にする
      scriptProperties.setProperty(propKey, cellValue);
      baseSheet.getRange(cell).setValue("設定済み");
      apiKeys.push(cellValue);
    }
  });

  if (apiKeys.length === 0) throw new Error('基礎データシートにAPIキーが設定されていません。');

  const baseUrl = `https://generativelanguage.googleapis.com/v1beta/models/${CONFIG.MODEL_NAME}:generateContent`;

  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { temperature: CONFIG.TEMPERATURE },
    safetySettings: [
      {
        category: "HARM_CATEGORY_HARASSMENT",
        threshold: "BLOCK_NONE"
      },
      {
        category: "HARM_CATEGORY_HATE_SPEECH",
        threshold: "BLOCK_NONE"
      },
      {
        category: "HARM_CATEGORY_SEXUALLY_EXPLICIT",
        threshold: "BLOCK_NONE"
      },
      {
        category: "HARM_CATEGORY_DANGEROUS_CONTENT",
        threshold: "BLOCK_NONE"
      }
    ]
  };

  if (isJson) payload.generationConfig.responseMimeType = "application/json";

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const maxRetries = 3;
  let lastError = null;

  for (let i = 0; i < apiKeys.length; i++) {
    const url = `${baseUrl}?key=${apiKeys[i]}`;
    let waitTime = 2000; // 初期待機時間：2秒

    for (let attempt = 0; attempt < maxRetries; attempt++) {
      try {
        const response = UrlFetchApp.fetch(url, options);
        const responseCode = response.getResponseCode();
        const responseText = response.getContentText();

        if (responseCode === 200) {
          const json = JSON.parse(responseText);
          if (json.candidates && json.candidates.length > 0) {
            const candidate = json.candidates[0];
            
            // 安全フィルターによるブロック判定
            if (candidate.finishReason === "SAFETY" || candidate.finishReason === "RECITATION") {
              throw new Error(`AI処理がブロックされました（理由: ${candidate.finishReason}）。安全フィルターまたは重複テキスト判定を調整してください。`);
            }
            
            if (candidate.content && candidate.content.parts && candidate.content.parts.length > 0) {
              return candidate.content.parts[0].text;
            }
          }
          throw new Error('AIから想定外の応答がありました。応答候補（candidates）が見つかりません。');
        } else if (responseCode === 429 || responseCode === 500 || responseCode === 503) {
          // レート制限または一時的なサーバーエラーの場合は、指数バックオフで待機してリトライ
          if (attempt < maxRetries - 1) {
            const sleepTime = waitTime + Math.random() * 1000;
            Utilities.sleep(sleepTime);
            waitTime *= 2;
            continue;
          } else {
            throw new Error(`API接続エラー(${responseCode}): リトライ上限に達しました。一時的に制限されているか、サーバーエラーです。`);
          }
        } else {
          // 400（キー不正）や 403（権限なし）などの恒久的なエラーはリトライせず、即座に次のAPIキーを試す
          throw new Error(`API接続エラー(${responseCode}): APIキーが無効であるか、不正なリクエストです。詳細: ${responseText}`);
        }
      } catch (e) {
        lastError = e;
        // リトライループを抜けて次のAPIキーに移る
        break; 
      }
    }
  }

  // すべてのAPIキーが失敗した場合に最後のエラーを投げる
  throw lastError || new Error('APIの呼び出しに失敗しました。');
}

/**
 * 【追加】AIが返した文字列からMarkdown装飾などを取り除き、安全にJSONオブジェクトに変換する
 */
function safeParseJSON(aiResponseText) {
  try {
    const cleanedText = aiResponseText.replace(/^\x60\x60\x60(?:json)?|\x60\x60\x60$/gm, '').trim();
    const parsed = JSON.parse(cleanedText);
    if (parsed && typeof parsed === 'object') {
      return parsed;
    }
    throw new Error('パースされたデータが有効なJSONオブジェクトではありません。');
  } catch (e) {
    throw new Error('AIの出力形式が不正でデータの読み取りに失敗しました。詳細: ' + e.message);
  }
}

function scanForPersonalInfo(sheet, startRow, lastRow, isReportCard, isYouroku, checkColStart, checkColEnd) {
  let warningRows = [];
  for (let r = startRow; r <= lastRow; r++) {
    let inputText = "";
    if (isReportCard) {
      for (let c = checkColStart; c <= checkColEnd; c++) {
        const val = sheet.getRange(r, c).getValue();
        if (val) inputText += val;
      }
    } else if (isYouroku) {
      inputText += sheet.getRange(r, CONFIG.COL_Y_TERM1).getValue() || "";
      inputText += sheet.getRange(r, CONFIG.COL_Y_TERM2).getValue() || "";
      inputText += sheet.getRange(r, CONFIG.COL_Y_TERM3).getValue() || "";
    } else {
      inputText += sheet.getRange(r, checkColStart).getValue() || "";
    }

    let testText = inputText;
    CONFIG.SAFE_WORDS.forEach(word => {
      testText = testText.split(word).join('');
    });
    if (CONFIG.REGEX_HONORIFICS.test(testText)) {
      warningRows.push(r);
    }
  }
  return warningRows;
}