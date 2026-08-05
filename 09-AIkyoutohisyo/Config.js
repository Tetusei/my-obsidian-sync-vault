/**
 * Config.gs
 * v3.1.7: 初期設定時のToDoシート入力規則再適用処理（ensureTodoSheetColumns）をループ処理から一括範囲設定（バッチ）に変更し、タイムアウトとフリーズ問題を修正
 */

const Config = {
  VERSION: 'v3.1.7', 
  FISCAL_YEAR: '2026年（令和8年）',
  
  SHEET_NAME_TODO: 'ToDo',
  SHEET_NAME_MASTER: '基礎データ',
  SHEET_NAME_DOC_FACTORY: '文書作成', 
  SHEET_NAME_DASHBOARD: 'ダッシュボード',
  
  BACKUP_FOLDER_NAME: '📦_バックアップ保存箱',
  REPORT_FOLDER_NAME: '日報',
  DRAFT_FOLDER_NAME: '文書作成', 
  MINUTES_FOLDER_NAME: '議事録_清書版',
  
  MASTER_POS: {
    WEBHOOK_CELL: 'B1',            // 個人用（朝の通知・ToDo含む）
    FORM_URL_CELL: 'B2',
    API_KEY_CELL: 'B3',
    TARGET_EMAIL_CELL: 'B4',
    MODEL_NAME_CELL: 'B5',
    CALENDAR_ID_CELL: 'B6',        // 全体用予定カレンダー
    REPORT_FOLDER_URL_CELL: 'B8',
    REPORT_ON_OFF_CELL: 'B9',
    OTHER_CALENDAR_ID_CELL: 'B10',
    ALL_WEBHOOK_CELL: 'B11',       // 全体用（リマインド等）
    REMINDER_ON_OFF_CELL: 'B12',
    MINUTES_MEMO_FOLDER_URL_CELL: 'B13', 
    ROSTER_FILE_URL_CELL: 'B14',  
    WORK_FOLDER_URL_CELL: 'B15',
    DRAFT_FOLDER_URL_CELL: 'B16',        
    MINUTES_SAVE_FOLDER_URL_CELL: 'B17', 
    TODO_CALENDAR_ID_CELL: 'B18',  // 新規：自分専用ToDoカレンダーID
    ATTACHMENTS_FOLDER_URL_CELL: 'B19',  // 新規：メール添付ファイル保存用フォルダ
    ROSTER_COL: 4
  },

  DOC_FACTORY_POS: {
    TITLE: 'B1', TARGET: 'B2', CONTENT: 'B3', TONE: 'B4', RESULT_URL: 'B6', SCHOOL_NAME: 'B7', PRINCIPAL_NAME: 'B8'
  },
  
  TODO_COL: {
    DATE: 0, SOURCE: 1, TITLE: 2, CONTENT: 3, PIC: 4, DUE_DATE: 5, PRIORITY: 6, STATUS: 7, 
    ACTION: 8, STAKEHOLDER: 9, NAME: 10, MAIL_LINK: 11, MEMO: 12, COMPLETED_DATE: 13
  },

  FORM_COL: {
    TIMESTAMP: 0, EMAIL: 1, TITLE: 2, CONTENT: 3, PIC: 4, DUE_DATE: 5, PRIORITY: 6, NEXT_ACTION: 7, MEMO: 8
  },
  FORM_QUESTION_TITLE_PIC: '担当',
  
  FORWARD_DESTINATIONS: {
    '進路指導部': 'shinro@example.com',
    '教務部': 'kyomu@example.com',
    '生徒指導部': 'seito@example.com',
    '事務室': 'jimu@example.com',
    '手元で留める（転送不要）': ''
  }
};

/**
 * スクリプトプロパティから安全にAPIキーを取得する（未格納時はセルから取得して格納）
 */
function getApiKey() {
  try {
    const propKey = "GEMINI_API_KEY";
    const props = PropertiesService.getScriptProperties();
    const savedKey = props.getProperty(propKey);
    if (savedKey) {
      return savedKey;
    }
    
    // プロパティにない場合、スプレッドシートのセルから読み取る
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(Config.SHEET_NAME_MASTER);
    if (!sheet) return "";
    
    const cell = sheet.getRange(Config.MASTER_POS.API_KEY_CELL);
    const val = String(cell.getValue() || "").trim();
    
    // マスク文字列以外で値が入っていれば保存してマスクする
    const skipValues = ["（設定済み）", "（格納済み）", "１本格納しました", "1本格納しました"];
    const isMasked = skipValues.some(v => val.startsWith(v));
    if (val && !isMasked) {
      props.setProperty(propKey, val);
      cell.setValue("（設定済み）");
      SpreadsheetApp.flush();
      return val;
    }
    return val;
  } catch(e) {
    console.error("getApiKey error:", e);
    return "";
  }
}

/**
 * スクリプトプロパティからすべてのAPIキーを取得する（未格納時はセルから取得して格納）
 * @return {string[]} APIキーの配列
 */
function getApiKeys() {
  try {
    const propKeys = "GEMINI_API_KEYS";
    const props = PropertiesService.getScriptProperties();
    const savedKeys = props.getProperty(propKeys);
    if (savedKeys) {
      const keys = JSON.parse(savedKeys);
      if (Array.isArray(keys) && keys.length > 0) {
        return keys;
      }
    }
    
    // 互換性のため単一キーを取得
    const singleKey = getApiKey();
    if (singleKey) {
      return [singleKey];
    }
    return [];
  } catch(e) {
    console.error("getApiKeys error:", e);
    return [];
  }
}

/**
 * 複数APIキーを自動ローテーションしてGemini APIを実行する
 * 一時的エラー（429や500、503など）が発生した場合は即座に次のキーで再試行します。
 * @param {Object} payload リクエストボディ
 * @param {string} modelName モデル名
 * @return {UrlFetchApp.HTTPResponse} レスポンスオブジェクト
 */
function callGeminiWithRotation(payload, modelName) {
  const apiKeys = getApiKeys();
  if (apiKeys.length === 0) {
    throw new Error("Gemini APIキーが設定されていません。スプレッドシートの「基礎データ」シートのB3セルに入力してください。");
  }
  
  const cacheKey = "GEMINI_KEY_INDEX";
  const cache = CacheService.getScriptCache();
  let currentIndex = 0;
  
  const cachedVal = cache.get(cacheKey);
  if (cachedVal !== null) {
    currentIndex = parseInt(cachedVal, 10) % apiKeys.length;
  }
  
  let lastError = null;
  const maxAttempts = Math.max(apiKeys.length * 2, 4);
  
  for (let attempt = 0; attempt < maxAttempts; attempt++) {
    const apiKey = apiKeys[currentIndex];
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${apiKey}`;
    
    const options = {
      "method": "post",
      "contentType": "application/json",
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true
    };
    
    try {
      const response = UrlFetchApp.fetch(url, options);
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();
      
      if (responseCode === 200) {
        // 成功した場合、次のキーインデックスをキャッシュに保存して終了
        const nextIndex = (currentIndex + 1) % apiKeys.length;
        cache.put(cacheKey, String(nextIndex), 21600); // 6時間キャッシュ
        return response;
      }
      
      // 一時的エラーまたはレート制限の場合は警告を出力して次のキーへ切り替え
      console.warn(`[APIキー切り替え警告] キー番号 ${currentIndex} がステータス ${responseCode} を返しました。次のキーで再試行します。レスポンス: ${responseText}`);
      lastError = `HTTP ${responseCode}: ${responseText}`;
      
      // ✨ 429 レート制限が発生した場合、APIが要求する時間または安全な時間スリープして自動回復を図る
      if (responseCode === 429) {
        let sleepMs = 15000; // デフォルトで15秒待機
        try {
          const errObj = JSON.parse(responseText);
          if (errObj.error && errObj.error.message && errObj.error.message.includes('Please retry in')) {
            // "Please retry in 35.102s" のようなメッセージから秒数をパース
            const match = errObj.error.message.match(/Please retry in ([\d.]+)\s*s/i);
            if (match) {
              sleepMs = Math.ceil(parseFloat(match[1]) * 1000) + 1500; // 1.5秒余裕を持たせる
            }
          }
        } catch (pErr) {}
        
        console.log(`[レート制限安全待機] 429エラーを検知したため、${Math.round(sleepMs / 1000)} 秒間スリープして自動復旧を待ちます...`);
        Utilities.sleep(sleepMs);
      }
      
    } catch (e) {
      console.warn(`[APIキー切り替え例外警告] キー番号 ${currentIndex} の呼び出し中に例外が発生しました: ${e.message}`);
      lastError = e.message;
    }
    
    // 次のインデックスに進む
    currentIndex = (currentIndex + 1) % apiKeys.length;
  }
  
  throw new Error(`すべてのAPIキーで呼び出しに失敗しました（試行回数: ${maxAttempts}）。最後のエラー: ${lastError}`);
}