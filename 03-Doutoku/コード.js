/**
 * 道徳所見生成スクリプト (v7.13 - 高速並列処理版・バックアップ機能付・完了通知追加・APIキー秘匿化・エラー対策・リトライ15回強化・タイムアウト対策・再実行スキップ・一括バックグラウンド作成・進捗状況表示追加)
 * 機能: メモなし対応、プロンプト切り替え対応、API安定化(fetchAll)、一括読み書きによる高速化
 */

const CONFIG = {
  SHEET_NAME_DATA: '基礎データ',
  SHEET_NAME_STUDENTS: '成長の様子',
  API_KEY_CELL: 'B1',
  PROMPT_CELL: 'B2',
  STATUS_CELL: 'B3', // 一括作成の実行状況表示セル

  COL_BASIC: 3,     // C列：基本文
  COL_OUTPUT: 4,    // D列：AI所見出力
  // E列(文字数)を飛ばす
  COL_CONTENT: 6,   // F列：内容項目
  COL_MATERIAL: 7,  // G列：題材名
  COL_MEMO: 9       // I列：振り返りメモ
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🌟道徳AI支援')
    .addItem('選択した行の所見を生成する', 'generateComments')
    .addItem('全生徒の所見を一括作成する', 'startGenerateAll')
    .addSeparator()
    .addItem('バックアップを作成する', 'createBackup')
    .addToUi();
}

function generateComments() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  // 1. 設定読み込み
  const keySheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME_DATA);
  if (!keySheet) {
    ui.alert(`エラー: '${CONFIG.SHEET_NAME_DATA}' シートが見つかりません。`);
    return;
  }
  
  let apiKeyValue = keySheet.getRange(CONFIG.API_KEY_CELL).getValue().toString().trim();
  const customPrompt = keySheet.getRange(CONFIG.PROMPT_CELL).getValue();

  // キーの登録・読み込み判定
  if (apiKeyValue && !apiKeyValue.startsWith('登録済み') && !apiKeyValue.startsWith('【登録済み】')) {
    // 生のAPIキーが入力されている場合、秘密のプロパティ領域に保存してセルをマスク
    const keys = apiKeyValue.split(/[\s,，、\n\r]+/).filter(k => k.length > 20);
    if (keys.length > 0) {
      PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEYS', JSON.stringify(keys));
      PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', keys[0]);
      keySheet.getRange(CONFIG.API_KEY_CELL).setValue(`登録済み (${keys.length}本のキー)`);
      SpreadsheetApp.flush();
    }
  }

  // APIキー配列の読み込み
  let apiKeys = [];
  try {
    const keysStr = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEYS');
    if (keysStr) apiKeys = JSON.parse(keysStr);
  } catch(e) {}
  
  if (apiKeys.length === 0) {
    const singleKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || PropertiesService.getUserProperties().getProperty('GEMINI_API_KEY');
    if (singleKey) apiKeys = [singleKey];
  }

  if (apiKeys.length === 0 || apiKeys[0].length < 20) {
    ui.alert(`エラー: APIキーが登録されていません。'${CONFIG.SHEET_NAME_DATA}' シートのB1セルにAPIキーを入力してください。`);
    return;
  }
  
  // プロンプトがエラー表示（#N/A等）や空欄の場合のガード
  if (!customPrompt || customPrompt === "【待機】モードを選択してください") {
    const response = ui.alert('確認', '基礎データシートで「モード」が選択されていないか、プロンプトが空です。\n標準設定（120文字）で実行しますか？', ui.ButtonSet.YES_NO);
    if (response !== ui.Button.YES) return;
  }

  // 2. 実行処理
  const selection = sheet.getSelection();
  const ranges = selection.getActiveRangeList().getRanges();

  spreadsheet.toast('所見の生成を開始します...', '🚀 スタート');

  // 全タスクのリスト
  const allTasks = [];

  // 各選択範囲のデータを一括で読み込み
  for (const range of ranges) {
    const startRow = range.getRow();
    const numRows = range.getNumRows();

    // 基本文(C列)からメモ(I列)まで一括取得
    const dataRange = sheet.getRange(startRow, CONFIG.COL_BASIC, numRows, CONFIG.COL_MEMO - CONFIG.COL_BASIC + 1);
    const dataValues = dataRange.getValues();

    // 出力用(D列)の一括範囲
    const outputRange = sheet.getRange(startRow, CONFIG.COL_OUTPUT, numRows, 1);
    const originalOutputs = outputRange.getValues();

    const rangeTasks = [];

    for (let i = 0; i < numRows; i++) {
      const rowIndex = startRow + i;
      const rowValues = dataValues[i];
      
      const basicText = rowValues[0]; // C列 (COL_BASIC)
      const content = rowValues[CONFIG.COL_CONTENT - CONFIG.COL_BASIC]; // F列
      const material = rowValues[CONFIG.COL_MATERIAL - CONFIG.COL_BASIC]; // G列
      const memo = rowValues[CONFIG.COL_MEMO - CONFIG.COL_BASIC]; // I列

      // 既存の出力内容を取得
      const currentOutput = originalOutputs[i][0] ? originalOutputs[i][0].toString().trim() : "";

      // スキップ判定
      // 1. メモも基本文も空ならスキップ
      // 2. すでに正常な文章（空ではなく、かつ「エラー:」や一時表示「（AIが」で始まらない）が入っている場合はスキップ
      const isAlreadyGenerated = currentOutput !== "" && 
                                  !currentOutput.startsWith("エラー:") && 
                                  !currentOutput.startsWith("（AIが");

      if ((!memo && !basicText) || isAlreadyGenerated) {
        rangeTasks.push({
          skip: true,
          originalOutput: originalOutputs[i][0]
        });
        continue;
      }

      const task = {
        skip: false,
        rowIndex: rowIndex,
        basicText: basicText,
        content: content,
        material: material,
        memo: memo,
        originalOutput: originalOutputs[i][0],
        success: false,
        result: ""
      };
      
      rangeTasks.push(task);
      allTasks.push(task);
    }

    range.associatedTasks = rangeTasks;
    range.outputRange = outputRange;
  }

  // 実行対象タスクがない場合
  if (allTasks.length === 0) {
    spreadsheet.toast('処理対象の行が見つかりませんでした。', '⚠️ スキップ', 3);
    return;
  }

  // ★時間制限対策：行数が多すぎる場合の警告アラート
  if (allTasks.length > 40) {
    const response = ui.alert(
      '確認',
      `一度に選択された行数（${allTasks.length}行）が多いため、Googleの実行制限時間（6分）を超える可能性があります。\n処理の遅延を防ぐため、1回につき30〜40行程度に分割して実行することをお勧めします。\n\nこのまま処理を実行しますか？`,
      ui.ButtonSet.YES_NO
    );
    if (response !== ui.Button.YES) return;
  }

  // 実行時間の監視開始
  const startTime = new Date().getTime();

  // 1. 進行状況「AIが執筆中...」を一括表示
  for (const range of ranges) {
    const loadingValues = range.associatedTasks.map(t => {
      return t.skip ? [t.originalOutput] : ["（AIが執筆中...🖊️）"];
    });
    range.outputRange.setValues(loadingValues);
  }
  SpreadsheetApp.flush(); // 画面を即座に更新

  // 2. 並列リクエストの構築
  let pendingTasks = allTasks.map((task, index) => {
    const keyIndex = index % apiKeys.length;
    const currentApiKey = apiKeys[keyIndex];
    const payload = buildGeminiPayload(task.basicText, task.content, task.material, task.memo, customPrompt);
    const model = 'gemini-2.5-flash';
    const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${currentApiKey}`;
    
    task.fetchParams = {
      url: apiUrl,
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    task.apiKeyIndex = keyIndex;
    return task;
  });

  // リトライ制御用
  let attempts = 0;
  const maxAttempts = 15; // 最大15回リトライするように強化
  let delay = 2000; // 初期待機時間 (2秒)

  while (pendingTasks.length > 0 && attempts < maxAttempts) {
    // ★時間制限対策：経過時間が4.5分（270,000ミリ秒）を超えた場合、安全にループを終了して書き戻す
    const elapsed = new Date().getTime() - startTime;
    if (elapsed > 270000) {
      for (const task of pendingTasks) {
        task.result = "エラー: 制限時間（6分）が近づいたため処理を一時中断しました。APIが一時的に混雑している可能性があるため、しばらく時間を置いてから、再度この行のみを選択して実行してください。";
        task.success = false;
      }
      break;
    }

    const urlsAndParams = pendingTasks.map(t => t.fetchParams);
    
    // fetchAllによる並列リクエストの送信
    const responses = UrlFetchApp.fetchAll(urlsAndParams);
    const nextPending = [];

    for (let k = 0; k < responses.length; k++) {
      const resp = responses[k];
      const task = pendingTasks[k];
      const statusCode = resp.getResponseCode();
      const contentText = resp.getContentText();

      if (statusCode === 200) {
        try {
          const json = JSON.parse(contentText);
          if (json.candidates && json.candidates[0] && json.candidates[0].content && json.candidates[0].content.parts && json.candidates[0].content.parts[0]) {
            task.result = json.candidates[0].content.parts[0].text.trim();
            task.success = true;
          } else {
            task.result = "エラー: 応答データ形式が正しくありません。";
            task.success = false;
          }
        } catch (e) {
          task.result = "エラー: " + e.message;
          task.success = false;
        }
      } else if (statusCode === 429 || statusCode === 502 || statusCode === 503 || statusCode === 504) {
        // 利用制限（429）、過負荷（503）、ゲートウェイエラー（502/504）はリトライ対象とする
        // 次回のリトライ時は、別のAPIキーを使うように更新する
        task.apiKeyIndex = (task.apiKeyIndex + 1) % apiKeys.length;
        const nextApiKey = apiKeys[task.apiKeyIndex];
        const model = 'gemini-2.5-flash';
        task.fetchParams.url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${nextApiKey}`;
        nextPending.push(task);
      } else {
        // その他のエラー
        try {
          const json = JSON.parse(contentText);
          task.result = "エラー: " + (json.error ? json.error.message : "HTTP " + statusCode);
        } catch (e) {
          task.result = "エラー: HTTP " + statusCode;
        }
        task.success = false;
      }
    }

    if (nextPending.length > 0) {
      attempts++;
      if (attempts < maxAttempts) {
        Utilities.sleep(delay);
        delay = Math.min(delay * 2, 10000); // 指数バックオフ（最大10秒に制限してタイムアウトを防ぐ）
        pendingTasks = nextPending;
      } else {
        // リトライ上限超過
        for (const task of nextPending) {
          task.result = "エラー: API利用上限（429 Too Many Requests）を超過しました。時間をおいて再試行してください。";
          task.success = false;
        }
      }
    } else {
      break;
    }
  }

  // 3. 結果を一括書き戻し
  for (const range of ranges) {
    const finalValues = range.associatedTasks.map(t => {
      return t.skip ? [t.originalOutput] : [t.result];
    });
    range.outputRange.setValues(finalValues);
  }

  spreadsheet.toast('すべての処理が完了しました！', '✅ 完了', 5);
}

// プロンプトおよびリクエストボディの構築関数
function buildGeminiPayload(basicText, content, material, memo, customInstruction) {
  const defaultInstruction = `
    文字数は120文字以内を厳守。
    基本文の構成を活かしつつ、メモの内容を組み込んでください。
    メモがない場合は、基本文を少し整える程度で出力してください。
  `;

  let activeInstruction = (customInstruction && customInstruction !== "【待機】モードを選択してください") 
                          ? customInstruction : defaultInstruction;

  let memoStatusInstruction = "";
  if (!memo || memo.trim() === "") {
    memoStatusInstruction = `
    【メモ欠落時の対応】
    児童の「振り返りメモ」は空欄です。「基本文」の内容だけを使用し、自然な評価文として出力してください。
    `;
  } else {
    memoStatusInstruction = `
    【作成方針】
    「基本文」をベースに、「振り返りメモ」の具体性を融合させてください。
    文字数オーバーを防ぐため、基本文の重複部分は大胆に削除・短縮してください。
    `;
  }

  const prompt = `
あなたは日本の学校教師です。道徳の授業における児童生徒の評価文を作成してください。

### 入力情報
- 内容項目: ${content}
- 教材名: ${material}
- ベースとなる基本文: ${basicText}
- 児童の振り返りメモ: ${memo ? memo : "(なし)"}

### 作成指示
${activeInstruction}

${memoStatusInstruction}

### 出力条件
- 生成された文章のみを出力してください。
  `;

  return {
    "contents": [{ "parts": [{ "text": prompt }] }]
  };
}

// 互換性維持のための単一呼び出し関数（内部リファクタリング）
function callGeminiAPI(apiKey, basicText, content, material, memo, customInstruction) {
  const model = 'gemini-2.5-flash';
  const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;

  const payload = buildGeminiPayload(basicText, content, material, memo, customInstruction);

  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  const response = UrlFetchApp.fetch(apiUrl, options);
  const json = JSON.parse(response.getContentText());

  if (json.error) throw new Error(json.error.message);
  return json.candidates[0].content.parts[0].text.trim();
}

// スプレッドシートと同じフォルダに「バックアップ」フォルダを作成し、コピーを作成する関数
function createBackup() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  try {
    const fileId = spreadsheet.getId();
    const file = DriveApp.getFileById(fileId);
    const parents = file.getParents();
    
    if (!parents.hasNext()) {
      ui.alert('エラー: スプレッドシートの親フォルダが見つかりません。');
      return;
    }
    
    const parentFolder = parents.next();
    const folderName = "バックアップ";
    
    // 「バックアップ」フォルダが存在するか確認、なければ作成
    const folders = parentFolder.getFoldersByName(folderName);
    let backupFolder;
    if (folders.hasNext()) {
      backupFolder = folders.next();
    } else {
      backupFolder = parentFolder.createFolder(folderName);
    }
    
    // 現在日時を取得してファイル名を作成
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
    const backupName = `${spreadsheet.getName()}_バックアップ_${timestamp}`;
    
    // バックアップを作成
    spreadsheet.toast('バックアップを作成しています...', '📁 処理中');
    file.makeCopy(backupName, backupFolder);
    spreadsheet.toast('バックアップを作成しました！', '✅ 完了', 5);
    ui.alert('完了', 'バックアップの作成が完了しました。', ui.ButtonSet.OK);
    
  } catch (e) {
    console.error(e);
    ui.alert('バックアップ作成エラー: ' + e.message);
  }
}

// セル編集時にAPIキーが入力された場合、自動で保存して表示をマスクするトリガー
function onEdit(e) {
  if (!e) return;
  try {
    const range = e.range;
    const sheet = range.getSheet();
    
    // 基礎データシートのB1セルが編集されたか判定
    if (sheet.getName() === CONFIG.SHEET_NAME_DATA && range.getRow() === 1 && range.getColumn() === 2) {
      const val = range.getValue().toString().trim();
      
      // 入力値が空でなく、「登録済み」で始まらない、十分な長さの値をAPIキーとして扱う
      if (val && !val.startsWith('登録済み') && !val.startsWith('【登録済み】') && val.length > 20) {
        // カンマ、改行、スペース等で分割してAPIキーを抽出
        const keys = val.split(/[\s,，、\n\r]+/).filter(k => k.length > 20);
        
        if (keys.length > 0) {
          PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEYS', JSON.stringify(keys));
          // 互換性のため、1番目のキーを単一プロパティにも保存
          PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', keys[0]);
          PropertiesService.getUserProperties().setProperty('GEMINI_API_KEY', keys[0]);
          
          range.setValue(`登録済み (${keys.length}本のキー)`);
          e.source.toast(`APIキーを${keys.length}本登録し、セル入力を保護しました。`, '🔑 登録完了');
        }
      }
    }
  } catch (err) {
    console.error("onEdit error:", err);
  }
}

// === 全生徒の一括バックグラウンド作成機能 ===

// メニューから実行されるトリガー設定・開始用関数
function startGenerateAll() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '一括作成の開始',
    'B列に氏名が入力されているすべての生徒の所見作成を開始します。\n\n' +
    '【動作の仕様】\n' +
    '・すでに正常に作成が完了している行は自動でスキップされます。\n' +
    '・実行時間制限（6分）にかかりそうになると自動で一時中断し、1分後に続きから自動で再開します。\n' +
    '・エラーが発生した行のみを抽出し、すべて完了するまでこの処理を自動で繰り返します。\n\n' +
    '実行を開始してもよろしいですか？',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;
  
  // 既存の古いトリガーがあれば削除
  deleteExistingTriggers();
  updateStatus(spreadsheet, "【実行中】処理を開始しました...");
  
  // 初回実行
  generateAllComments();
}

// 既存の一括生成トリガーを削除するヘルパー関数
function deleteExistingTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'generateAllComments') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
}

// 全生徒の所見を巡回作成する関数（時間監視付き）
function generateAllComments() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME_STUDENTS);
  
  if (!sheet) {
    spreadsheet.toast("エラー: '成長の様子' シートが見つかりません。", "❌ 中断");
    deleteExistingTriggers();
    updateStatus(spreadsheet, "【停止】エラー：'成長の様子' シートが見つかりません。");
    return;
  }
  
  // 1. 設定の取得と認証キーの読み込み
  const keySheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME_DATA);
  if (!keySheet) {
    spreadsheet.toast("エラー: '基礎データ' シートが見つかりません。", "❌ 中断");
    deleteExistingTriggers();
    updateStatus(spreadsheet, "【停止】エラー：基礎データシートが見つかりません。");
    return;
  }
  
  const customPrompt = keySheet.getRange(CONFIG.PROMPT_CELL).getValue();
  
  // APIキー配列の読み込み
  let apiKeys = [];
  try {
    const keysStr = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEYS');
    if (keysStr) apiKeys = JSON.parse(keysStr);
  } catch(e) {}
  
  if (apiKeys.length === 0) {
    const singleKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || PropertiesService.getUserProperties().getProperty('GEMINI_API_KEY');
    if (singleKey) apiKeys = [singleKey];
  }
  
  if (apiKeys.length === 0 || apiKeys[0].length < 20) {
    spreadsheet.toast("エラー: APIキーが登録されていません。基礎データシートのB1セルにAPIキーを入力してください。", "❌ 中断");
    deleteExistingTriggers();
    updateStatus(spreadsheet, "【停止】エラー：APIキーが登録されていません。B1セルにキーを入力してください。");
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 4) {
    spreadsheet.toast("処理対象の生徒データ（4行目以降）がありません。", "⚠️ 終了");
    deleteExistingTriggers();
    updateStatus(spreadsheet, "【停止】エラー：処理対象の生徒データがありません。");
    return;
  }

  // 4行目以降の全生徒データを一括取得 (B列からI列まで、幅8列)
  const startRow = 4;
  const numRows = lastRow - startRow + 1;
  const dataRange = sheet.getRange(startRow, 2, numRows, 8); 
  const dataValues = dataRange.getValues();
  
  const tasks = [];
  
  for (let i = 0; i < numRows; i++) {
    const rowIndex = startRow + i;
    const rowValues = dataValues[i];
    
    const name = rowValues[0]; // B列
    const basicText = rowValues[1]; // C列
    const currentOutput = rowValues[2] ? rowValues[2].toString().trim() : ""; // D列
    const content = rowValues[4]; // F列
    const material = rowValues[5]; // G列
    const memo = rowValues[7]; // I列
    
    // 生徒氏名が入力されている行のみを対象にする
    if (name && name.toString().trim() !== "") {
      // スキップ判定: すでに正常な所見（空でなく、かつ「エラー:」や一時表示「（AIが」で始まらないもの）が入っている場合はスキップ
      const isAlreadyGenerated = currentOutput !== "" && 
                                 !currentOutput.startsWith("エラー:") && 
                                 !currentOutput.startsWith("（AIが");
      
      if (!isAlreadyGenerated) {
        tasks.push({
          rowIndex: rowIndex,
          basicText: basicText,
          content: content,
          material: material,
          memo: memo,
          originalOutput: rowValues[2]
        });
      }
    }
  }
  
  // すべて完了している場合
  if (tasks.length === 0) {
    spreadsheet.toast('すべての生徒の所見作成が完了しました！', '✅ 完了', 10);
    deleteExistingTriggers();
    
    const now = new Date();
    const dateTimeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");
    updateStatus(spreadsheet, `【完了】すべての生徒の所見作成が完了しました！（完了日時: ${dateTimeStr}）`);
    
    try {
      SpreadsheetApp.getUi().alert('完了', 'すべての生徒の所見作成が完了しました！', SpreadsheetApp.getUi().ButtonSet.OK);
    } catch(e) {}
    return;
  }
  
  // バックグラウンドで進行中のトーストを表示
  spreadsheet.toast(`残りの ${tasks.length} 名の所見を作成しています...`, '⚙️ 処理進行中', 5);
  
  const now = new Date();
  const timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "HH:mm:ss");
  updateStatus(spreadsheet, `【実行中】残りの ${tasks.length} 名の所見を作成中...（最終更新: ${timeStr}）`);

  const startTime = new Date().getTime();
  let processedCount = 0;

  // キーのインデックス状態を保持するオブジェクト
  const keyState = { index: 0 };

  for (let i = 0; i < tasks.length; i++) {
    const task = tasks[i];

    // 経過時間のチェック（4分 = 240,000ミリ秒 を超えたら早期離脱して次回トリガーへ引き継ぐ）
    const elapsed = new Date().getTime() - startTime;
    if (elapsed > 240000) {
      deleteExistingTriggers();
      ScriptApp.newTrigger('generateAllComments')
        .timeBased()
        .after(120000) // API制限解除のため2分後に自動再開
        .create();
      
      const nextRun = new Date(new Date().getTime() + 120000);
      const nextTimeStr = Utilities.formatDate(nextRun, Session.getScriptTimeZone(), "HH:mm:ss");
      updateStatus(spreadsheet, `【待機中】一時停止しました。${nextTimeStr} 頃に自動再開します。（残り ${tasks.length - processedCount} 名）`);
      
      spreadsheet.toast(`時間制限のため一時停止しました。残り ${tasks.length - processedCount} 名の処理を2分後に自動再開します。`, '⏳ 一時待機', 10);
      return;
    }
    
    // 1. セルに「執筆中...」を書き込み
    sheet.getRange(task.rowIndex, CONFIG.COL_OUTPUT).setValue("（AIが執筆中...🖊️）");
    SpreadsheetApp.flush();
    
    // 2. 個別API呼び出しとリトライ（直列処理＋複数キー自動切り替え）
    processSingleTaskWithRetry(task, apiKeys, customPrompt, sheet, keyState);
    
    processedCount++;

    // APIの 15 RPM 制限（1分間に15回）を超えないようウェイトを設ける
    // 登録されたキー本数が多ければ、ウェイトを短縮または無しにする
    if (i < tasks.length - 1) {
      if (apiKeys.length === 1) {
        Utilities.sleep(4000); // 1本なら4秒待機
      } else if (apiKeys.length === 2) {
        Utilities.sleep(2000); // 2本なら2秒待機
      } else {
        Utilities.sleep(500);  // 3本以上なら0.5秒待機（ローテーションにより制限回避）
      }
    }
  }
  
  // 処理が一周完了した場合、再度自身を呼び出してエラー行の漏れがないか最終確認・再スキャンする
  deleteExistingTriggers();
  ScriptApp.newTrigger('generateAllComments')
    .timeBased()
    .after(120000) // API制限やエラー状態をリセットするため、再スキャンは2分後に実行
    .create();
    
  const nextCheck = new Date(new Date().getTime() + 120000);
  const nextCheckStr = Utilities.formatDate(nextCheck, Session.getScriptTimeZone(), "HH:mm:ss");
  updateStatus(spreadsheet, `【待機中】一周完了。再スキャンとエラー検証を ${nextCheckStr} 頃に実行します。（エラーがなければ完了となります）`);
}

// 単一のタスクを複数キーの切り替え・リトライ制御付きで実行する
function processSingleTaskWithRetry(task, apiKeys, customPrompt, sheet, keyState) {
  const model = 'gemini-2.5-flash';
  
  let attempts = 0;
  const maxAttempts = Math.max(apiKeys.length * 2, 4); // キー本数に応じた最大試行数
  let success = false;
  let resultText = "";

  while (attempts < maxAttempts) {
    // 現在のキーを取得
    const apiKey = apiKeys[keyState.index];
    const payload = buildGeminiPayload(task.basicText, task.content, task.material, task.memo, customPrompt);
    const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;
    
    const options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    try {
      const resp = UrlFetchApp.fetch(apiUrl, options);
      const statusCode = resp.getResponseCode();
      const contentText = resp.getContentText();

      if (statusCode === 200) {
        const json = JSON.parse(contentText);
        if (json.candidates && json.candidates[0] && json.candidates[0].content && json.candidates[0].content.parts && json.candidates[0].content.parts[0]) {
          resultText = json.candidates[0].content.parts[0].text.trim();
          success = true;
          
          // 成功したら次のキーに進める（ローテーション）
          keyState.index = (keyState.index + 1) % apiKeys.length;
          break;
        } else {
          resultText = "エラー: [停止] 応答データ形式不正";
          break;
        }
      } else if (statusCode === 429 || statusCode === 502 || statusCode === 503 || statusCode === 504) {
        // 一時的エラーの場合は、即座に次のキーに切り替えて再試行
        attempts++;
        keyState.index = (keyState.index + 1) % apiKeys.length;
        
        // 全てのキーを一周してもダメな場合は少し待つ
        if (attempts % apiKeys.length === 0 && attempts < maxAttempts) {
          Utilities.sleep(2000);
        }
      } else {
        // 400や403など恒久的なエラー
        try {
          const json = JSON.parse(contentText);
          resultText = "エラー: [停止] " + (json.error ? json.error.message : "HTTP " + statusCode);
        } catch (e) {
          resultText = "エラー: [停止] HTTP " + statusCode;
        }
        break;
      }
    } catch (e) {
      attempts++;
      keyState.index = (keyState.index + 1) % apiKeys.length;
      if (attempts % apiKeys.length === 0 && attempts < maxAttempts) {
        Utilities.sleep(2000);
      }
    }
  }

  if (!success && resultText === "") {
    resultText = "エラー: [再試行待ち] すべてのAPIキーで制限またはエラーが発生しました。";
  }

  // 結果をセルに書き込む
  sheet.getRange(task.rowIndex, CONFIG.COL_OUTPUT).setValue(resultText);
  SpreadsheetApp.flush();
}

// 基礎データシートのセルA3・B3に進捗ステータスを書き込むヘルパー関数
function updateStatus(spreadsheet, statusText) {
  try {
    const keySheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME_DATA);
    if (keySheet) {
      keySheet.getRange("A3").setValue("一括作成ステータス");
      keySheet.getRange(CONFIG.STATUS_CELL).setValue(statusText);
      SpreadsheetApp.flush();
    }
  } catch (e) {
    console.error("Status update error:", e);
  }
}