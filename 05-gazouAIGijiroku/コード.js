/**
 * 【No.2】議事録作成ツール（ハイブリッド入力版）
 * C列：フォルダURL（中身を全部読む）
 * D〜H列：ファイルURL（個別に追加で読む）
 * → I列に議事録を出力
 */

// ▼▼ 設定エリア ▼▼
const CONFIG_SHEET_NAME_NO2 = '基礎データ';
const API_KEY_CELL_NO2 = 'B1';
const FOLDER_ID_CELL_NO2 = 'B3'; // 議事録の保存先
const MAX_FILES_IN_FOLDER = 20;  // フォルダ読込時の安全装置
// ▲▲ 設定エリア終了 ▲▲

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('★議事録作成') 
    .addItem('【No.2】資料から議事録作成', 'createMeetingMinutes')
    .addToUi();
}

/**
 * セルが編集されたときに自動実行されるシンプルなトリガー
 */
function onEdit(e) {
  if (!e) return;
  const sheet = e.source.getActiveSheet();
  const range = e.range;
  
  // 基礎データシートの B1セル (APIキー) が編集された場合のみ処理
  if (sheet.getName() === CONFIG_SHEET_NAME_NO2 && range.getA1Notation() === API_KEY_CELL_NO2) {
    const val = range.getValue().toString().trim();
    if (val !== "" && val !== "1本格納しています。") {
      // 内部に保存してセルの表示を上書き
      PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', val);
      range.setValue("1本格納しています。");
    } else if (val === "") {
      // 空欄にされた場合は保存されているプロパティも削除
      PropertiesService.getScriptProperties().deleteProperty('GEMINI_API_KEY');
    }
  }
}

function createMeetingMinutes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  
  if (sheet.getName() === CONFIG_SHEET_NAME_NO2) {
    Browser.msgBox("エラー", "「基礎データ」シートでは実行できません。", Browser.Buttons.OK);
    return;
  }
  const configSheet = ss.getSheetByName(CONFIG_SHEET_NAME_NO2);
  if (!configSheet) {
    Browser.msgBox("エラー", `「${CONFIG_SHEET_NAME_NO2}」シートが見つかりません。`, Browser.Buttons.OK);
    return;
  }

  const activeRange = sheet.getActiveRange();
  if (!activeRange) {
    Browser.msgBox("エラー", "行が選択されていません。", Browser.Buttons.OK);
    return;
  }
  const activeRow = activeRange.getRow();
  if (activeRow < 2) {
    Browser.msgBox("エラー", "データ行を選択してください。", Browser.Buttons.OK);
    return;
  }
  
  // 保存先フォルダ取得
  const outputFolderRaw = configSheet.getRange(FOLDER_ID_CELL_NO2).getValue();
  let outputFolder;
  try {
    outputFolder = DriveApp.getFolderById(extractIdForVision(outputFolderRaw));
  } catch(e) {
    Browser.msgBox("設定エラー", "保存先フォルダ(B3)が見つかりません。", Browser.Buttons.OK);
    return;
  }

  // ■■■ 資料収集ロジック ■■■
  let targetFiles = []; // {file: File, source: string}
  
  // 1. C列（フォルダ）のチェック
  const folderUrl = sheet.getRange(activeRow, 3).getValue();
  if (folderUrl && folderUrl.toString().trim() !== "") {
    try {
      const fId = extractIdForVision(folderUrl);
      const folder = DriveApp.getFolderById(fId);
      const filesIterator = folder.getFiles();
      let count = 0;
      while (filesIterator.hasNext()) {
        if (count >= MAX_FILES_IN_FOLDER) break;
        targetFiles.push({ file: filesIterator.next(), source: "C列フォルダ内" });
        count++;
      }
    } catch(e) {
      Browser.msgBox("フォルダエラー", `C列のフォルダを読み込めませんでした。\n${e.message}`, Browser.Buttons.OK);
      return;
    }
  }

  // 2. D〜H列（個別ファイル）のチェック
  // 4列目(D)から5列分(D,E,F,G,H)を取得
  const fileUrls = sheet.getRange(activeRow, 4, 1, 5).getValues()[0];
  const colNames = ["D", "E", "F", "G", "H"];
  
  fileUrls.forEach((url, i) => {
    if (url && url.toString().trim() !== "") {
      try {
        const fId = extractIdForVision(url);
        const file = DriveApp.getFileById(fId);
        targetFiles.push({ file: file, source: `${colNames[i]}列指定` });
      } catch(e) {
        // 個別指定のエラーは致命的なので停止して通知
        Browser.msgBox("ファイルエラー", `${colNames[i]}列のURLが無効です。\n${e.message}`, Browser.Buttons.OK);
        throw new Error("Invalid File URL"); 
      }
    }
  });

  // 何もなければ終了
  if (targetFiles.length === 0) {
    Browser.msgBox("エラー", "読み込む資料がありません。\nC列にフォルダ、またはD〜H列にファイルのURLを入力してください。", Browser.Buttons.OK);
    return;
  }

  // ■■■ 資料解析処理 ■■■
  const fileBlobs = [];   // 画像・PDF用
  let extractedText = ""; // テキスト用
  const validFiles = [];  // リスト作成用

  for (const item of targetFiles) {
    const file = item.file;
    const mime = file.getMimeType();
    const name = file.getName();
    
    try {
      if (mime === MimeType.GOOGLE_DOCS) {
        // ドキュメント
        const text = DocumentApp.openById(file.getId()).getBody().getText();
        extractedText += `\n\n--- 資料: ${name} (Doc) ---\n${text}\n`;
        validFiles.push(file);

      } else if (mime === MimeType.GOOGLE_SHEETS) {
        // スプレッドシート
        const ssObj = SpreadsheetApp.openById(file.getId());
        extractedText += `\n\n--- 資料: ${name} (Sheet) ---\n`;
        ssObj.getSheets().forEach(s => {
          if (s.getLastRow() > 0) {
            const vals = s.getDataRange().getValues();
            const csv = vals.map(r => r.join(" , ")).join("\n");
            extractedText += `\n[シート: ${s.getName()}]\n${csv}\n`;
          }
        });
        validFiles.push(file);

      } else if (mime.startsWith('image/') || mime === MimeType.PDF) {
        // 画像・PDF
        fileBlobs.push(file.getBlob());
        validFiles.push(file);

      } else {
        console.warn(`未対応形式スキップ: ${name}`);
      }
    } catch(e) {
      console.error(`読込失敗: ${name}`);
    }
  }

  // APIキーの取得（PropertiesServiceから取得、または未格納の場合はセルから取得）
  const apiKeyCell = configSheet.getRange(API_KEY_CELL_NO2);
  let apiKeyValue = apiKeyCell.getValue().toString().trim();
  let apiKey = "";

  if (apiKeyValue === "1本格納しています。") {
    apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  } else if (apiKeyValue !== "") {
    // onEditが未実行の場合などのセーフティ
    apiKey = apiKeyValue;
    PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', apiKey);
    apiKeyCell.setValue("1本格納しています。");
    SpreadsheetApp.flush();
  } else {
    // セルが空の場合、登録済みのキーを使用
    apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (apiKey) {
      apiKeyCell.setValue("1本格納しています。");
      SpreadsheetApp.flush();
    }
  }

  if (!apiKey) {
    Browser.msgBox("エラー", "APIキーが設定されていません。B1セルにAPIキーを入力してEnterキーを押してください。", Browser.Buttons.OK);
    return;
  }

  const date = sheet.getRange(activeRow, 1).getDisplayValue(); 
  const title = sheet.getRange(activeRow, 2).getValue();       
  
  ss.toast(`資料計${validFiles.length}点を分析中...`, 'AI処理中', 20);

  // プロンプト
  const docTitle = `${date}_${title}_議事録`;
  const prompt = `
  あなたは優秀な書記です。提供された資料（画像、PDF、テキスト、表データ）をすべて統合し、議事録を作成してください。
  
  【基本情報】
  日付: ${date}
  件名: ${title}

  【テキスト系資料の内容】
  ${extractedText}

  【指示】
  1. すべての資料の内容を網羅・統合して分析してください。
  2. 以下の形式で出力してください。
     - ■ 概要 (要約)
     - ■ 決定事項・Todo
     - ■ 詳細内容 (議論の流れ、板書、数値データなど)
  `;

  try {
    const generatedText = callGeminiUnified(apiKey, prompt, fileBlobs);

    const doc = DocumentApp.create(docTitle);
    const body = doc.getBody();
    body.insertParagraph(0, docTitle).setHeading(DocumentApp.ParagraphHeading.HEADING1);
    
    // 参照資料リスト
    body.appendParagraph("【参照資料リスト】");
    validFiles.forEach(file => {
      const name = file.getName();
      const url = file.getUrl();
      const mime = file.getMimeType();
      
      if (mime.startsWith('image/')) {
        try {
          body.appendParagraph(`[画像: ${name}]`);
          body.appendImage(file.getBlob()).setWidth(400);
        } catch(e) {
          body.appendParagraph(`[リンク: ${name}]`).setLinkUrl(url);
        }
      } else {
        body.appendParagraph(`[リンク: ${name}]`).setLinkUrl(url);
      }
      body.appendParagraph(""); 
    });
    
    body.appendParagraph("\n----------------\n");
    body.appendParagraph(generatedText);
    
    // 保存と出力 (I列 = 9列目)
    const resFile = DriveApp.getFileById(doc.getId());
    resFile.moveTo(outputFolder);
    sheet.getRange(activeRow, 9).setValue(doc.getUrl()); 
    
    Browser.msgBox("完了", `議事録を作成しました！\nI列をご確認ください。`, Browser.Buttons.OK);

  } catch (e) {
    Browser.msgBox("エラー発生", e.toString(), Browser.Buttons.OK);
    console.error(e);
  }
}

// Gemini API呼び出し
function callGeminiUnified(apiKey, prompt, fileBlobs) {
  const modelName = 'gemini-flash-latest';
  const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${apiKey}`;
  const parts = [{ "text": prompt }];
  if (fileBlobs && fileBlobs.length > 0) {
    fileBlobs.forEach(blob => {
      parts.push({
        "inline_data": {
          "mime_type": blob.getContentType(),
          "data": Utilities.base64Encode(blob.getBytes())
        }
      });
    });
  }
  const payload = { "contents": [{ "parts": parts }] };
  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };
  const response = UrlFetchApp.fetch(apiUrl, options);
  const json = JSON.parse(response.getContentText());
  if (json.error) throw new Error(`API Error: ${json.error.message}`);
  if (json.candidates && json.candidates.length > 0 && json.candidates[0].content) {
    return json.candidates[0].content.parts[0].text;
  } else {
    throw new Error("AIからの応答が空でした。");
  }
}

// ID抽出ヘルパー
function extractIdForVision(input) {
  if (!input) return null;
  const text = input.toString().trim();
  if (text.includes("/d/")) return text.split("/d/")[1].split(/[/?]/)[0];
  if (text.includes("folders/")) return text.split("folders/")[1].split(/[/?]/)[0];
  if (text.includes("id=")) return text.split("id=")[1].split("&")[0];
  return text;
}