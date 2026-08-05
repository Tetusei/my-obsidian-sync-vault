/**
 * PhoneMemoOps.gs (電話伝言自動仕分け)
 * 音声フォーム送信時に起動し、AIが伝言先を判別してチャット通知とToDo登録を行う関数
 * ※ 教職員名簿シートのF列（役職）の読み込みに対応
 */
function processPhoneMemo(e) {
  // 1. スプレッドシートのイベントオブジェクトからデータを取得
  if (!e || !e.namedValues) {
    console.log("フォーム送信イベントが正しく取得できませんでした。手動実行はできません。");
    return;
  }

  let memoContent = "";
  for (let key in e.namedValues) {
    if (key.indexOf("メモ") !== -1 || key === "メモ内容") {
      memoContent = e.namedValues[key][0];
      break;
    }
  }
  
  if (!memoContent) {
    console.log("メモ内容が空のため処理をスキップしました。");
    return;
  }

  // 2. 基礎データシートから各種設定を取得
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(Config.SHEET_NAME_MASTER);
  const apiKey = getApiKey();
  const modelName = masterSheet.getRange(Config.MASTER_POS.MODEL_NAME_CELL).getValue() || "gemini-2.5-flash";
  const chatWebhookUrl = masterSheet.getRange(Config.MASTER_POS.ALL_WEBHOOK_CELL).getValue(); 
  const rosterUrl = masterSheet.getRange(Config.MASTER_POS.ROSTER_FILE_URL_CELL).getValue(); 

  if (!apiKey) {
    console.log("APIキーが設定されていません。");
    return;
  }

  // 3. 名簿データの取得（★F列の役職対応に改修）
  let rosterText = "【教職員名簿（名前・役職・担当備考）】\n";
  try {
    if (rosterUrl) {
      const rosterSs = SpreadsheetApp.openByUrl(rosterUrl);
      const staffSheet = rosterSs.getSheetByName("教職員") || rosterSs.getSheets()[0];
      const staffData = staffSheet.getDataRange().getValues();
      
      for (let r = 1; r < staffData.length; r++) {
        const row = staffData[r];
        const name = row[0];       // A列: 氏名
        const memo = row[1] || ""; // B列: 備考
        // F列（インデックス5）から役職を取得（列が足りない場合のエラーを安全に回避）
        const role = (row.length > 5) ? (row[5] || "") : ""; 
        
        // AIに「氏名」「役職」「備考」をすべて分かりやすく伝えます
        rosterText += `- ${name}先生 (役職: ${role || "なし"}, 備考: ${memo || "なし"})\n`;
      }
      
      const studentSheet = rosterSs.getSheetByName("生徒");
      if (studentSheet) {
        rosterText += "\n【生徒名簿（年組と氏名）】\n";
        const studentData = studentSheet.getDataRange().getValues();
        for (let s = 1; s < studentData.length; s++) {
          rosterText += "- " + studentData[s][0] + "年" + studentData[s][1] + "組: " + studentData[s][2] + "\n";
        }
      }
    }
  } catch(err) {
    console.log("名簿読み込み失敗: " + err);
    rosterText += "(名簿ファイルが一時的に読み込めませんでした。文脈から伝言先を予測してください)\n";
  }

  // 4. Gemini APIに仕分けを依頼するプロンプトの作成
  const apiURL = `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${apiKey}`;
  
  const prompt = `あなたは学校の優秀な教頭秘書AIです。以下の【電話・伝言メモ】を読み、提供された【名簿データ】を参考に、誰宛ての伝言かを特定してください。

【電話・伝言メモ】
${memoContent}

${rosterText}

【命令】
1. 伝言の宛先となる教職員の「氏名（苗字のみで可）」を特定してください。もし役職（例：教頭、保健室、1年担任など）しか分からない場合は、名簿の役職欄や備考欄から該当する先生を特定してください。どうしても特定できない場合は「全員」としてください。
2. ToDoシートに登録するための「分かりやすい件名（20文字程度）」を作成してください。
3. メモ内容をきれいに整理した「内容（詳細）」を作成してください。
4. 職員チャットに投稿するための、丁寧な「チャット用メッセージ」を作成してください（宛先の先生の名前を冒頭に含める）。

【出力フォーマット】
必ず以下の項目を持つ純粋なJSON形式（装飾なし）のみで出力してください。余計な説明文は一切不要です。
{
  "targetStaff": "特定された先生の苗字、または全員",
  "subject": "【伝言】〇〇に関する件 などの件名",
  "body": "要約・整理された詳細内容",
  "chatMessage": "チャット用の通知文章"
}`;

  const payload = {
    "contents": [{ "parts": [{ "text": prompt }] }],
    "generationConfig": { "responseMimeType": "application/json" }
  };

  // 5. APIを呼び出して結果を解析
  let resultJson;
  try {
    const response = callGeminiWithRotation(payload, modelName);
    const resultText = JSON.parse(response.getContentText()).candidates[0].content.parts[0].text;
    resultJson = JSON.parse(resultText.trim());
  } catch (e) {
    console.log("AI解析エラー: " + e.message);
    resultJson = {
      targetStaff: "確認待ち",
      subject: "【AI解析失敗】電話伝言",
      body: memoContent,
      chatMessage: `【AI自動解析エラー】\n以下の伝言が届いていますが仕分けに失敗しました。教頭先生、確認をお願いします。\n\n${memoContent}`
    };
  }

  // 6. ToDoシートへの自動登録
  const todoSheet = ss.getSheetByName(Config.SHEET_NAME_TODO);
  if (todoSheet) {
    const today = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd");
    const newRow = [];
    newRow[Config.TODO_COL.DATE] = today;
    newRow[Config.TODO_COL.SOURCE] = "電話メモ";
    newRow[Config.TODO_COL.TITLE] = resultJson.subject;
    newRow[Config.TODO_COL.CONTENT] = resultJson.body;
    newRow[Config.TODO_COL.PIC] = "教頭";
    newRow[Config.TODO_COL.DUE_DATE] = "";
    newRow[Config.TODO_COL.PRIORITY] = "中";
    newRow[Config.TODO_COL.STATUS] = "未着手";
    newRow[Config.TODO_COL.ACTION] = false;
    newRow[Config.TODO_COL.STAKEHOLDER] = resultJson.targetStaff;
    newRow[Config.TODO_COL.MAIL_LINK] = "";
    newRow[Config.TODO_COL.MEMO] = resultJson.targetStaff ? `【電話伝言】宛先: ${resultJson.targetStaff}先生` : "【電話伝言】全員宛";
    newRow[Config.TODO_COL.COMPLETED_DATE] = "";
    
    todoSheet.appendRow(newRow);
    const lastRow = todoSheet.getLastRow();
    todoSheet.getRange(lastRow, Config.TODO_COL.ACTION + 1).insertCheckboxes();
    setTodoRowValidations(todoSheet, lastRow);
  }

  // 7. 職員チャット（Webhook）への自動通知
  if (chatWebhookUrl && resultJson.chatMessage) {
    try {
      const chatOptions = {
        "method": "post",
        "contentType": "application/json",
        "payload": JSON.stringify({ "text": resultJson.chatMessage })
      };
      UrlFetchApp.fetch(chatWebhookUrl, chatOptions);
    } catch(e) {
      console.log("チャット送信エラー: " + e.message);
    }
  }
}