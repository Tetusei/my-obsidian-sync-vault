/**
 * 【GASレシピ Final 100-Point Version】
 * * 修正点:
 * 1. 配点を「合計100点」になるよう調整しました。
 * - 5問の場合（選択3:記述2）→ 選択12点 / 記述32点
 * - 10問の場合（選択7:記述3）→ 選択7点 / 記述17点
 * 2. その他の全機能（AIモデル、重複回避、A列リンク転記など）は全て維持。
 */

// ▼▼▼ 設定エリア ▼▼▼
const CONFIG = {
  MODEL_NAME: "gemini-flash-latest", 
  POINTS: 20, 
  SHEET_DATA: '基礎データ',
  SHEET_MEMBER: 'クラス名簿'
};
// ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('★テスト業務')
    .addItem('1. AIでテストを作成する', 'showSettingDialog')
    .addSeparator()
    .addItem('2. 成績を名簿に転記する', 'importScoresToMemberSheet')
    .addToUi();
}

/**
 * 1. 設定用ダイアログ
 */
function showSettingDialog() {
  const htmlString = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body { font-family: sans-serif; padding: 10px; font-size: 14px; }
          .group { margin-bottom: 15px; border-bottom: 1px solid #eee; padding-bottom: 10px; }
          label { display: block; margin-bottom: 5px; font-weight: bold; color: #333; }
          .radio-group label { font-weight: normal; margin-right: 15px; display: inline; cursor: pointer; }
          .checkbox-group { margin-top: 5px; }
          .checkbox-group div { margin-bottom: 5px; }
          button { background-color: #4285f4; color: white; border: none; padding: 12px 20px; border-radius: 4px; cursor: pointer; font-size: 14px; width: 100%; margin-top: 5px; }
          button:hover { background-color: #357ae8; }
          button:disabled { background-color: #ccc; cursor: not-allowed; }
        </style>
      </head>
      <body>
        <div class="group">
          <label>1. 問題数</label>
          <div class="radio-group">
            <input type="radio" id="q5" name="qCount" value="5" checked>
            <label for="q5">5問</label>
            <input type="radio" id="q10" name="qCount" value="10">
            <label for="q10">10問</label>
          </div>
        </div>

        <div class="group">
          <label>2. 難易度</label>
          <div class="radio-group">
            <input type="radio" id="easy" name="difficulty" value="やさしい">
            <label for="easy">やさしい</label>
            <input type="radio" id="normal" name="difficulty" value="普通" checked>
            <label for="normal">普通</label>
            <input type="radio" id="hard" name="difficulty" value="難しい">
            <label for="hard">難しい</label>
          </div>
        </div>

        <div class="group" style="border-bottom: none;">
          <label>3. オプション</label>
          <div class="checkbox-group">
            <div>
              <input type="checkbox" id="includeText" name="includeText" checked>
              <label for="includeText" style="font-weight:normal;">記述式問題を含める</label>
            </div>
            <div>
              <input type="checkbox" id="shuffle" name="shuffle">
              <label for="shuffle" style="font-weight:normal;">問題をシャッフルする（カンニング防止）</label>
            </div>
          </div>
        </div>

        <button id="btn" onclick="runScript()">作成開始</button>
        <script>
          function runScript() {
            const btn = document.getElementById('btn');
            btn.disabled = true;
            btn.textContent = '準備中...ウィンドウが閉じたら処理が始まります';
            
            const count = document.querySelector('input[name="qCount"]:checked').value;
            const difficulty = document.querySelector('input[name="difficulty"]:checked').value;
            const hasText = document.getElementById('includeText').checked;
            const isShuffle = document.getElementById('shuffle').checked;
            
            google.script.run
              .withSuccessHandler(function() { google.script.host.close(); })
              .withFailureHandler(function(e) { alert('エラー: ' + e); btn.disabled = false; })
              .startQuizCreation({ 
                count: Number(count), 
                difficulty: difficulty, 
                hasText: hasText,
                isShuffle: isShuffle
              });
          }
        </script>
      </body>
    </html>
  `;
  const html = HtmlService.createHtmlOutput(htmlString).setWidth(330).setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'テスト作成設定');
}

/**
 * 2. テスト作成メイン処理
 */
function startQuizCreation(options) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_DATA);
  const memberSheet = ss.getSheetByName(CONFIG.SHEET_MEMBER);

  if (!sheet || !memberSheet) { ui.alert("エラー: 必要なシートが見つかりません。"); return; }

  let apiKey = sheet.getRange("B1").getValue().toString().trim();
  if (apiKey === "１本格納しました" || apiKey === "1本格納しました" || apiKey === "1本格納しています。") {
    apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || "";
  }
  const docUrl = sheet.getRange("B3").getValue();

  if (!apiKey || !docUrl) { ui.alert("エラー: APIキーまたはドキュメントURLが入力されていません。"); return; }

  ss.toast("ドキュメントを読み込んでいます...", "処理開始", 5);

  try {
    let doc;
    try { doc = DocumentApp.openByUrl(docUrl); } 
    catch (e) { ui.alert("エラー: ドキュメントを開けませんでした。URLを確認してください。"); return; }

    const docBody = doc.getBody();
    const textContent = docBody.getText();
    const images = getImagesFromDoc(docBody); 

    if (textContent.replace(/\s/g, "").length < 10 && images.length === 0) {
      ui.alert("エラー: ドキュメントに十分なテキストも画像も見当たりません。"); return;
    }

    ui.alert(`「${doc.getName()}」の内容からテストを作成します。\n（難易度: ${options.difficulty} / シャッフル: ${options.isShuffle ? "あり" : "なし"}）\n\nAIが問題を生成しています。このまま1分ほどお待ちください...`);

    // AIへリクエスト
    const quizData = callGeminiAPI(apiKey, textContent, images, options);
    if (!quizData || quizData.length === 0) throw new Error("AIからの応答から問題を生成できませんでした。");

    // 番号（No）決定ロジック
    const bValues = sheet.getRange("B:B").getValues();
    let lastRow = 0;
    for (let i = bValues.length - 1; i >= 0; i--) {
      if (bValues[i][0] !== "") {
        lastRow = i + 1;
        break;
      }
    }
    if (lastRow < 5) lastRow = 5; 
    const targetRow = lastRow + 1;
    
    let historyNo = 1; 
    if (targetRow > 6) { 
      const prevNoCell = sheet.getRange(targetRow - 1, 1);
      const prevNo = prevNoCell.getValue();
      if (typeof prevNo === 'number') historyNo = prevNo + 1;
      else historyNo = targetRow - 5; 
    }

    // ファイル名・日時決定（重複チェック機能付き）
    const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd");
    const baseDocName = doc.getName();
    
    // 既存のタイトル一覧を取得
    const existingTitles = sheet.getRange("C:C").getValues().flat();
    
    let tempDocName = baseDocName;
    let formTitle = `${tempDocName} ${dateStr} (${options.difficulty})`;
    let counter = 2;

    // 重複していたら「名前②」「名前③」と数字を増やしていく
    while (existingTitles.includes(formTitle)) {
      const circleNums = ["", "", "②", "③", "④", "⑤", "⑥", "⑦", "⑧", "⑨", "⑩"];
      const suffix = circleNums[counter] || `(${counter})`; 
      
      tempDocName = `${baseDocName}${suffix}`;
      formTitle = `${tempDocName} ${dateStr} (${options.difficulty})`;
      counter++;
    }

    // フォーム作成
    const formInfo = createGoogleForm(formTitle, quizData, ss.getId(), images, options.isShuffle);
    let responseSheetId = null;

    // シート操作
    try {
      SpreadsheetApp.flush(); 
      const allSheets = ss.getSheets();
      
      for (const s of allSheets) {
        const sUrl = s.getFormUrl();
        if (sUrl && sUrl.indexOf(formInfo.id) !== -1) {
          
          let newName = String(historyNo);
          let renameCounter = 2;
          while (ss.getSheetByName(newName)) {
            newName = `${historyNo}_${renameCounter}`;
            renameCounter++;
          }
          s.setName(newName); 
          responseSheetId = s.getSheetId();
          
          ss.setActiveSheet(s);
          ss.moveActiveSheet(ss.getNumSheets()); 
          break;
        }
      }

      const sheetDescription = ss.getSheetByName('説明');
      if(sheetDescription) {
        ss.setActiveSheet(sheetDescription);
        ss.moveActiveSheet(1);
      }
      ss.setActiveSheet(sheet);
      ss.moveActiveSheet(2);
      ss.setActiveSheet(memberSheet); 
      ss.moveActiveSheet(3);

      ss.setActiveSheet(sheet);
      sheet.getRange("A1").activate();

    } catch (e) { console.log("シート操作エラー: " + e.message); }

    // 名簿ヘッダー書き込み
    const lastCol = memberSheet.getLastColumn();
    const targetCol = lastCol + 1;
    const lastMemberRow = memberSheet.getLastRow();
    const headerCell = memberSheet.getRange(1, targetCol);
    
    // タイトルから安全な名前を作る
    const safeDocName = tempDocName.replace(/"/g, '""'); 
    const formula = `=HYPERLINK("${formInfo.url}", "${safeDocName}" & CHAR(10) & "${dateStr}")`;
    
    headerCell.setFormula(formula); 
    headerCell.setWrap(true); 
    headerCell.setVerticalAlignment("middle");
    headerCell.setHorizontalAlignment("center");
    
    if (lastMemberRow > 1) {
      memberSheet.getRange(1, targetCol, lastMemberRow, 1).setBorder(true, true, true, true, true, true);
    }
    
    // 履歴保存（基礎データ）
    if (responseSheetId !== null) {
      sheet.getRange(targetRow, 1).setFormula(`=HYPERLINK("#gid=${responseSheetId}", "${historyNo}")`);
    } else {
      sheet.getRange(targetRow, 1).setValue(historyNo);
    }

    sheet.getRange(targetRow, 2).setValue(new Date());         
    sheet.getRange(targetRow, 3).setValue(formTitle); // 重複回避済みのタイトル       
    sheet.getRange(targetRow, 4).setValue(formInfo.url);       
    sheet.getRange(targetRow, 5).setValue(formInfo.publishedUrl); 

    // QRコード（リンク版）
    const qrUrl = "https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=";
    const qrFormula = `=HYPERLINK("${qrUrl}" & ENCODEURL(E${targetRow}), "🔗QR表示")`;
    sheet.getRange(targetRow, 6).setFormula(qrFormula);

    // 書式設定
    sheet.getRange(targetRow, 1).setHorizontalAlignment("right");
    sheet.getRange(targetRow, 4, 1, 2).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
    sheet.getRange(targetRow, 6).setHorizontalAlignment("center");
    sheet.getRange(targetRow, 1, 1, 6).setVerticalAlignment("middle");
    sheet.getRange(targetRow, 1, 1, 6).setBorder(true, true, true, true, true, true);

    showSuccessDialog(formInfo.url);

  } catch (e) {
    console.error(e);
    ui.alert(`エラーが発生しました:\n${e.message}`);
  }
}

/**
 * 完了ダイアログ
 */
function showSuccessDialog(url) {
  const htmlTemplate = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body { font-family: sans-serif; text-align: center; padding: 20px; }
          .msg { margin-bottom: 20px; font-size: 16px; color: #333; font-weight: bold; }
          .btn { 
            display: inline-block; background-color: #4CAF50; color: white; 
            padding: 12px 24px; text-decoration: none; border-radius: 4px; 
            font-weight: bold; font-size: 16px; 
          }
          .btn:hover { background-color: #45a049; }
          .alert-box {
            margin-top: 20px; padding: 10px; background-color: #fff3cd; 
            border: 1px solid #ffeeba; border-radius: 4px; color: #856404; font-size: 12px; text-align: left;
          }
          .sub { margin-top: 15px; font-size: 12px; color: #666; }
        </style>
        <script>
          window.onload = function() {
            var win = window.open("${url}", "_blank");
          };
        </script>
      </head>
      <body>
        <div class="msg">✨ 作成が完了しました！</div>
        <div>
          <a href="${url}" target="_blank" class="btn">フォームを開く</a>
        </div>
        
        <div class="alert-box">
          <strong>⚠️ メールアドレスが「手入力」の場合</strong><br>
          学校等の制限により自動設定ができませんでした。<br>
          お手数ですが、フォームの<strong>[設定]タブ → [回答]</strong> にある<br>
          「メールアドレスを収集する」を<strong>「確認済み」</strong>に変更してください。
        </div>

        <div class="sub">
          ※基礎データF列に「QR表示リンク」を作成しました。
        </div>
        <br>
        <button onclick="google.script.host.close()">閉じる</button>
      </body>
    </html>
  `;
  const html = HtmlService.createHtmlOutput(htmlTemplate).setWidth(400).setHeight(420);
  SpreadsheetApp.getUi().showModalDialog(html, '処理完了');
}

/**
 * 3. 成績転記処理（A列リンク活用版 + 平均点行クリーンアップ）
 */
function importScoresToMemberSheet() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_DATA);
  const memberSheet = ss.getSheetByName(CONFIG.SHEET_MEMBER);

  const bValues = sheet.getRange("B:B").getValues();
  let lastRow = 0;
  for (let i = bValues.length - 1; i >= 0; i--) {
    if (bValues[i][0] !== "") {
      lastRow = i + 1;
      break;
    }
  }

  if (lastRow < 6) { ui.alert("まだテストが作成されていません。"); return; }
  
  const range = sheet.getRange(6, 1, lastRow - 5, 3);
  const dataValues = range.getValues();
  const dataFormulas = range.getFormulas();

  // ★平均点行の重複チェックとクリーンアップロジック★
  const memberLastRow = memberSheet.getLastRow();
  const rangeA = memberSheet.getRange(1, 1, memberLastRow, 1);
  const valuesA = rangeA.getValues();
  let avgRow = -1;
  const rowsToDelete = [];

  for (let i = 0; i < valuesA.length; i++) {
    if (String(valuesA[i][0]) === "平均点") {
      if (avgRow === -1) {
        avgRow = i + 1; 
      } else {
        rowsToDelete.push(i + 1); 
      }
    }
  }
  
  if (rowsToDelete.length > 0) {
    rowsToDelete.sort((a, b) => b - a);
    rowsToDelete.forEach(row => memberSheet.deleteRow(row));
  }

  if (avgRow === -1) {
    avgRow = memberSheet.getLastRow() + 1;
    memberSheet.getRange(avgRow, 1).setValue("平均点").setFontWeight("bold").setHorizontalAlignment("right");
  }

  const lastStudentRow = avgRow - 1;
  if (lastStudentRow < 2) { ui.alert("生徒データが見つかりません。"); return; }

  const memberData = memberSheet.getRange(1, 1, avgRow, memberSheet.getLastColumn()).getValues();
  const header = memberData[0]; 
  const members = memberData.slice(1, lastStudentRow); 
  
  let emailColIndex = header.indexOf("メールアドレス");
  if (emailColIndex === -1) emailColIndex = 4; 

  let totalUpdateCount = 0;
  let allSheets = ss.getSheets(); 

  for (let i = 0; i < dataValues.length; i++) {
    const title = dataValues[i][2]; 
    const formula = dataFormulas[i][0]; 
    
    if (!title || !formula) continue;

    const match = formula.match(/gid=(\d+)/);
    if (!match) continue;
    const targetGid = Number(match[1]);

    const targetSheet = allSheets.find(s => s.getSheetId() === targetGid);
    if (!targetSheet) continue; 

    const responseData = targetSheet.getDataRange().getValues();
    if (responseData.length < 2) continue; 

    const resHeader = responseData[0];
    let resEmailIdx = resHeader.indexOf("メールアドレス");
    let resScoreIdx = resHeader.indexOf("スコア");
    if (resEmailIdx === -1 || resScoreIdx === -1) continue;

    const cleanTitle = String(title).replace(/[\s\u3000\(\)（）普通易しい難しい]/g, "");
    
    const destColIndex = header.findIndex(h => {
      const cleanHeader = String(h).replace(/[\s\u3000\n]/g, "");
      return cleanHeader !== "" && cleanTitle.indexOf(cleanHeader) !== -1;
    });

    if (destColIndex === -1) continue; 

    let updateCount = 0;
    
    for (let r = 1; r < responseData.length; r++) {
      const email = responseData[r][resEmailIdx];
      const scoreStr = responseData[r][resScoreIdx]; 

      let score = 0;
      if (typeof scoreStr === 'number') {
        score = scoreStr;
      } else {
        const scoreMatch = String(scoreStr).match(/^(\d+)/);
        if (scoreMatch) score = Number(scoreMatch[1]);
      }

      for (let m = 0; m < members.length; m++) {
        if (members[m][emailColIndex] === email) {
          memberSheet.getRange(m + 2, destColIndex + 1).setValue(score);
          updateCount++;
          break;
        }
      }
    }
    
    if (updateCount > 0) {
      totalUpdateCount += updateCount;
      
      const colLetter = columnToLetter(destColIndex + 1);
      const avgFormula = `=AVERAGE(${colLetter}2:${colLetter}${lastStudentRow})`;
      const cell = memberSheet.getRange(avgRow, destColIndex + 1);
      cell.setFormula(avgFormula);
      cell.setNumberFormat("0.0");
      cell.setFontWeight("bold");
    }
  }
  
  ui.alert(`転記完了！\n合計 ${totalUpdateCount} 件の点数を更新し、重複していた平均点行は整理されました。`);
}

function columnToLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function callGeminiAPI(key, text, images, options) {
  const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/${CONFIG.MODEL_NAME}:generateContent?key=${key}`;

  let mcCount = options.count;
  let textCount = 0;
  if (options.hasText) {
    if (options.count === 5) { mcCount = 3; textCount = 2; } 
    else if (options.count === 10) { mcCount = 7; textCount = 3; }
  }

  let difficultyPrompt = "";
  if (options.difficulty === "やさしい") {
    difficultyPrompt = "問題の難易度は「やさしい」にしてください。基礎的な用語の確認や、本文に明確な答えがある問題を優先し、ひっかけ問題は避けてください。";
  } else if (options.difficulty === "難しい") {
    difficultyPrompt = "問題の難易度は「難しい」にしてください。単なる暗記ではなく、文脈の理解や応用力を問う問題にしてください。選択肢にはもっともらしい誤答（distractors）を含めてください。";
  } else {
    difficultyPrompt = "問題の難易度は「普通」にしてください。標準的な理解度を確認できるレベルにしてください。";
  }

  const promptText = `
あなたは学校の先生です。提示された資料（テキストおよび画像）の内容を読み取り、生徒の理解度を確認するテストを作成してください。

【作成条件】
1. 合計で **${options.count}問** 作成してください。
   - 4択問題（正解は1つ）: ${mcCount}問
   - 記述式問題（短答）: ${textCount}問
2. **${difficultyPrompt}**
3. 画像に関する問題を作成する場合、問題文には「資料画像Xを参照して...」のように、画像の番号を明記してください。
4. 出力は必ず以下のJSON形式のみで行い、Markdown記法は含めないでください。

【JSONフォーマット】
[
  {
    "type": "multiple_choice",
    "question": "問題文...",
    "options": ["選択肢1", "選択肢2", "選択肢3", "選択肢4"],
    "answer": "正解の選択肢",
    "explanation": "解説..."
  },
  {
    "type": "text",
    "question": "記述問題文...",
    "explanation": "模範解答..."
  }
]

【教材テキスト】
${text}
`;

  const parts = [{ text: promptText }];
  images.forEach(img => {
    parts.push({ inline_data: { mime_type: img.mimeType, data: img.data } });
  });

  const payload = {
    contents: [{ parts: parts }],
    generationConfig: { responseMimeType: "application/json" }
  };

  const optionsHttp = {
    method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true
  };

  const maxRetries = 3;
  for (let i = 0; i < maxRetries; i++) {
    try {
      const response = UrlFetchApp.fetch(endpoint, optionsHttp);
      const json = JSON.parse(response.getContentText());
      if (!json.error) return JSON.parse(json.candidates[0].content.parts[0].text);
      
      console.warn(`Attempt ${i+1} Error: ${json.error.message}`);
      
      SpreadsheetApp.getActiveSpreadsheet().toast("AIへのアクセスが集中しています。80秒間待機します...", "制限回避モード", 80);
      SpreadsheetApp.flush(); 
      Utilities.sleep(80000); 
      
      if (i === maxRetries - 1) throw new Error(`AI API Error: ${json.error.message}`);
    } catch (e) {
      if (i === maxRetries - 1) throw e;
      SpreadsheetApp.getActiveSpreadsheet().toast("通信エラー。80秒後に再試行します...", "待機中", 80);
      SpreadsheetApp.flush();
      Utilities.sleep(80000); 
    }
  }
}

/**
 * フォーム作成（★修正：配点100点確定版）
 */
function createGoogleForm(title, questions, targetFolderFileId, images, isShuffle) {
  const form = FormApp.create(title);
  form.setIsQuiz(true);
  
  form.setCollectEmail(true);

  try {
    form.setRequireLogin(true); 
  } catch (e) {
    console.log("制限設定のエラー: " + e.message);
  }

  form.setDestination(FormApp.DestinationType.SPREADSHEET, targetFolderFileId);
  form.setDescription("AI自動生成テスト");

  if (isShuffle) {
    form.setShuffleQuestions(true);
  }

  if (images && images.length > 0) {
    const imgHeader = form.addSectionHeaderItem();
    imgHeader.setTitle("【参考資料】画像を参考に下の問題に答えてください");
    images.forEach((img, i) => form.addImageItem().setImage(img.blob).setTitle(`資料画像 ${i + 1}`));
    form.addSectionHeaderItem().setTitle("▼▼▼ ここから問題です ▼▼▼");
  }
  
  // ▼▼▼ 配点計算ロジック（100点ジャストにする） ▼▼▼
  let mcCount = 0;
  let textCount = 0;
  questions.forEach(q => {
    if(q.type === 'multiple_choice') mcCount++;
    else textCount++;
  });

  let mcPoint = 0;
  let textPoint = 0;

  // 5問（選択3:記述2）の標準ケース → 選択12点/記述32点
  if (mcCount === 3 && textCount === 2) {
    mcPoint = 12;   // 12 * 3 = 36
    textPoint = 32; // 32 * 2 = 64 -> 合計100点
  }
  // 10問（選択7:記述3）の標準ケース → 選択7点/記述17点
  else if (mcCount === 7 && textCount === 3) {
    mcPoint = 7;    // 7 * 7 = 49
    textPoint = 17; // 17 * 3 = 51 -> 合計100点
  }
  // それ以外のイレギュラーな構成の場合（汎用計算）
  else {
    let basePoint = 0;
    if (textCount === 0) {
      basePoint = 100 / mcCount;
      mcPoint = Math.floor(basePoint);
    } else if (mcCount === 0) {
      basePoint = 100 / textCount;
      textPoint = Math.floor(basePoint);
    } else {
      // 記述を2倍重み付けで計算
      basePoint = 100 / (mcCount + (textCount * 2));
      mcPoint = Math.floor(basePoint);
      textPoint = Math.floor(basePoint * 2);
    }
  }
  // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

  questions.forEach((q, i) => {
    const point = (q.type === 'multiple_choice') ? mcPoint : textPoint;

    if (q.type === 'multiple_choice') {
      const item = form.addMultipleChoiceItem();
      item.setTitle(`問${i+1}. ${q.question}`).setPoints(point).setRequired(true);
      item.setChoices(q.options.map(opt => item.createChoice(opt, opt === q.answer)));
      if (q.explanation) item.setFeedbackForCorrect(FormApp.createFeedback().setText(q.explanation).build());
    } else {
      const item = form.addParagraphTextItem();
      item.setTitle(`問${i+1}. ${q.question}`).setPoints(point).setRequired(true);
      if (q.explanation) item.setGeneralFeedback(FormApp.createFeedback().setText(`【模範解答】\n${q.explanation}`).build());
    }
  });

  try {
    const targetFile = DriveApp.getFileById(targetFolderFileId); 
    const formFile = DriveApp.getFileById(form.getId()); 
    formFile.moveTo(targetFile.getParents().next()); 
  } catch(e) {}

  return { 
    id: form.getId(), 
    url: form.getEditUrl(), 
    publishedUrl: form.getPublishedUrl() 
  };
}

function getImagesFromDoc(body) {
  const images = [];
  const numChildren = body.getNumChildren();
  for (let i = 0; i < numChildren; i++) {
    const child = body.getChild(i);
    if (child.getType() === DocumentApp.ElementType.PARAGRAPH) processContainer(child.asParagraph(), images);
    else if (child.getType() === DocumentApp.ElementType.LIST_ITEM) processContainer(child.asListItem(), images);
    if (images.length >= 5) break; 
  }
  return images;
}

function processContainer(container, images) {
  for (let j = 0; j < container.getNumChildren(); j++) {
    const element = container.getChild(j);
    if (element.getType() === DocumentApp.ElementType.INLINE_IMAGE) {
      const imgBlob = element.asInlineImage().getBlob();
      if (imgBlob.getContentType().startsWith('image/')) {
        images.push({ blob: imgBlob, mimeType: imgBlob.getContentType(), data: Utilities.base64Encode(imgBlob.getBytes()) });
      }
    }
  }
}

/**
 * セルが編集されたときに自動実行されるシンプルなトリガー
 * B1セルに入力されたAPIキーをPropertiesServiceに退避し、表示をマスクします。
 */
function onEdit(e) {
  if (!e) return;
  try {
    const sheet = e.source.getActiveSheet();
    const range = e.range;
    
    if (sheet.getName() === CONFIG.SHEET_DATA && range.getA1Notation() === 'B1') {
      const val = range.getValue().toString().trim();
      if (val !== "" && val !== "１本格納しました" && val !== "1本格納しました") {
        PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', val);
        range.setValue("１本格納しました");
        e.source.toast("APIキーを安全に格納しました。", "🔑 格納完了");
      } else if (val === "") {
        PropertiesService.getScriptProperties().deleteProperty('GEMINI_API_KEY');
        e.source.toast("APIキーを削除しました。", "🗑️ 削除完了");
      }
    }
  } catch (err) {
    console.error("onEdit error:", err);
  }
}