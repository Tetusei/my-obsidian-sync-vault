/**
 * 【スクリプト名】アンケート自由記述分析＆回答案作成ツール (Ver.4.0 チャット搭載版)
 * 【概要】
 * 外部シートをGeminiで分析し、以下の機能を提供します。
 * 1. 分析レポート（スライド・グラフ）作成
 * 2. 回答案（ドキュメント）作成
 * 3. ★新機能: AIデータ分析チャット（サイドバー）
 *
 * * 【設定（基礎データシート）】
 * B1: Gemini APIキー
 * B3: 分析したいスプレッドシートのURL
 * B4: 分析したいシート名
 * B5: 分析対象の列アルファベット
 * B6: 追加の指示
 *
 * * 【作成者】ICT支援員 (GAS作成)
 * 【更新日】2026/01/13
 */

// ==========================================
// 設定エリア
// ==========================================

const GEMINI_MODEL = 'gemini-2.5-flash'; 
const CONFIG_SHEET_NAME = '基礎データ';
const RESULT_SHEET_NAME = '分析結果';

// ==========================================
// メイン処理（メニュー作成）
// ==========================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚡ アンケート分析')
    .addItem('📊 1. 分析実行（グラフ・スライド作成）', 'generateAnalysisSlides')
    .addSeparator()
    .addItem('📝 2. 回答案作成（Googleドキュメント）', 'generateResponseDoc')
    .addSeparator()
    .addItem('💬 3. AIに質問（データ分析チャット）', 'showChatSidebar') // ★新メニュー
    .addToUi();
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
    
    if (sheet.getName() === CONFIG_SHEET_NAME && range.getA1Notation() === 'B1') {
      const val = range.getValue().toString().trim();
      if (val !== "" && val !== "1本格納しています。") {
        PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', val);
        range.setValue("1本格納しています。");
        e.source.toast("APIキーを安全に格納しました。", "🔑 格納完了");
      } else if (val === "") {
        PropertiesService.getScriptProperties().deleteProperty('GEMINI_API_KEY');
      }
    }
  } catch (err) {
    console.error("onEdit error:", err);
  }
}

// ==========================================
// 機能1：分析＆スライド作成
// ==========================================

function generateAnalysisSlides() {
  const ui = SpreadsheetApp.getUi();
  const currentSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const config = getConfig(currentSpreadsheet);
  if (!config) return;

  try {
    const dataObj = getTargetData(config);
    if (!dataObj) return;

    currentSpreadsheet.toast('AIが全体傾向とカテゴリ分類を分析中...', 'AI分析中', -1);
    const analysisResult = analyzeWithGemini(dataObj.comments, config.apiKey, config.instruction);

    if (!analysisResult) { ui.alert('分析に失敗しました。'); return; }

    currentSpreadsheet.toast('グラフを作成しています...', 'グラフ作成', -1);
    const sentimentChartBlob = createSentimentChartBlob(analysisResult); 
    const keywordChartBlob = createKeywordChartBlob(analysisResult);    

    currentSpreadsheet.toast('レポートスライドを作成中...', '仕上げ', -1);
    const parentFolder = DriveApp.getFileById(config.targetSpreadsheetId).getParents().next();
    const slideUrl = createSlidePresentation(analysisResult, dataObj.fileName, parentFolder, sentimentChartBlob, keywordChartBlob, config.instruction);

    saveResultToSheet(currentSpreadsheet, analysisResult, dataObj.fileName, config.sheetName, config.url, slideUrl, config.instruction);
    showCompleteModal(ui, slideUrl, 'スライド', RESULT_SHEET_NAME);

  } catch (e) {
    console.error(e);
    ui.alert('エラー発生', `処理を中断しました。\n詳細: ${e.toString()}`, ui.ButtonSet.OK);
  }
}

// ==========================================
// 機能2：保護者向け回答案作成
// ==========================================

function generateResponseDoc() {
  const ui = SpreadsheetApp.getUi();
  const currentSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const config = getConfig(currentSpreadsheet);
  if (!config) return;

  try {
    const dataObj = getTargetData(config);
    if (!dataObj) return;

    currentSpreadsheet.toast('回答案を構成し、表データを作成しています...', 'AI執筆中', -1);
    const docData = generateDocContentWithGemini(dataObj.comments, config.apiKey, config.instruction);

    if (!docData) { ui.alert('作成に失敗しました。'); return; }

    currentSpreadsheet.toast('Googleドキュメントに書き出し中（表組み整形）...', 'ファイル作成', -1);
    const parentFolder = DriveApp.getFileById(config.targetSpreadsheetId).getParents().next();
    const docUrl = createResponseDoc(docData, dataObj.fileName, parentFolder, config.instruction);

    const htmlOutput = HtmlService.createHtmlOutput(`<p>文書の作成が完了しました！</p><ul><li>表組みがきれいに整形されて挿入されました。</li></ul><p><strong>内容を確認し、加筆・修正してご利用ください。</strong></p><p><a href="${docUrl}" target="_blank">作成されたドキュメントを開く</a></p>`).setWidth(400).setHeight(200);
    ui.showModalDialog(htmlOutput, '作成完了');

  } catch (e) {
    console.error(e);
    ui.alert('エラー発生', `処理を中断しました。\n詳細: ${e.toString()}`, ui.ButtonSet.OK);
  }
}

// ==========================================
// 機能3：AIデータ分析チャット（★新機能）
// ==========================================

function showChatSidebar() {
  const htmlContent = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&display=swap" rel="stylesheet">
        <style>
          * { box-sizing: border-box; }
          body { 
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; 
            padding: 12px; 
            display: flex; 
            flex-direction: column; 
            height: 98vh; 
            color: #1f2937; 
            background-color: #f9fafb;
            margin: 0;
          }
          .header {
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 8px;
            padding-bottom: 12px;
            border-bottom: 1px solid #e5e7eb;
            margin-bottom: 12px;
          }
          h3 { 
            margin: 0; 
            color: #1e3a8a; 
            font-size: 15px; 
            font-weight: 600;
          }
          #chat-history { 
            flex-grow: 1; 
            overflow-y: auto; 
            padding: 12px; 
            margin-bottom: 12px; 
            border-radius: 12px; 
            background: #ffffff; 
            border: 1px solid #e5e7eb;
            box-shadow: inset 0 2px 4px rgba(0,0,0,0.02);
            display: flex;
            flex-direction: column;
            gap: 12px;
          }
          .message { 
            padding: 10px 14px; 
            border-radius: 14px; 
            max-width: 88%; 
            line-height: 1.5; 
            font-size: 13px; 
            word-wrap: break-word;
            animation: fadeIn 0.2s ease-out;
          }
          .user { 
            background-color: #3b82f6; 
            color: white; 
            align-self: flex-end; 
            border-bottom-right-radius: 2px; 
            box-shadow: 0 2px 4px rgba(59, 130, 246, 0.2);
          }
          .ai { 
            background-color: #f3f4f6; 
            color: #1f2937; 
            align-self: flex-start; 
            border-bottom-left-radius: 2px; 
            border: 1px solid #e5e7eb;
          }
          .ai ul, .ai ol {
            margin: 4px 0;
            padding-left: 20px;
          }
          .ai li {
            margin-bottom: 4px;
          }
          .ai strong {
            color: #1e3a8a;
            font-weight: 600;
          }
          .input-area { 
            display: flex; 
            gap: 8px; 
            align-items: center;
          }
          input[type="text"] { 
            flex-grow: 1; 
            padding: 10px 14px; 
            border-radius: 24px; 
            border: 1px solid #d1d5db; 
            outline: none; 
            font-size: 13px;
            background: #ffffff;
            transition: all 0.2s ease;
            box-shadow: 0 1px 2px rgba(0,0,0,0.05);
          }
          input[type="text"]:focus { 
            border-color: #3b82f6; 
            box-shadow: 0 0 0 3px rgba(59, 130, 246, 0.15);
          }
          button { 
            height: 38px;
            width: 38px;
            min-width: 38px;
            background-color: #3b82f6; 
            color: white; 
            border: none; 
            border-radius: 50%; 
            cursor: pointer; 
            display: flex;
            align-items: center;
            justify-content: center;
            transition: all 0.2s ease;
            box-shadow: 0 2px 4px rgba(59, 130, 246, 0.2);
          }
          button:disabled { 
            background-color: #d1d5db; 
            cursor: default; 
            box-shadow: none;
          }
          button:hover:not(:disabled) { 
            background-color: #2563eb; 
            transform: scale(1.05);
          }
          button:active:not(:disabled) { 
            transform: scale(0.95);
          }
          /* Typing indicator animation */
          .typing-indicator {
            display: flex;
            align-items: center;
            gap: 4px;
            padding: 10px 14px;
            background: #f3f4f6;
            border: 1px solid #e5e7eb;
            border-radius: 14px;
            border-bottom-left-radius: 2px;
            align-self: flex-start;
            margin-right: auto;
            display: none;
          }
          .typing-indicator span {
            width: 6px;
            height: 6px;
            background: #9ca3af;
            border-radius: 50%;
            display: inline-block;
            animation: bounce 1.4s infinite ease-in-out both;
          }
          .typing-indicator span:nth-child(1) { animation-delay: -0.32s; }
          .typing-indicator span:nth-child(2) { animation-delay: -0.16s; }
          .typing-indicator span:nth-child(3) { animation-delay: 0s; }
          
          @keyframes bounce {
            0%, 80%, 100% { transform: scale(0); }
            40% { transform: scale(1.0); }
          }
          @keyframes fadeIn {
            from { opacity: 0; transform: translateY(5px); }
            to { opacity: 1; transform: translateY(0); }
          }
          .example-chips {
            font-size: 11px;
            color: #6b7280;
            margin-top: 6px;
            line-height: 1.4;
          }
        </style>
      </head>
      <body>
        <div class="header">
          <span>🤖</span>
          <h3>アンケート分析チャット</h3>
        </div>
        <div id="chat-history">
          <div class="message ai">
            こんにちは！このアンケートデータについて、何でも質問してください。
            <div class="example-chips">
              <strong>質問例:</strong><br>
              ・給食についての意見はある？<br>
              ・1年生の保護者は何を心配してる？<br>
              ・「ありがとう」という言葉は含まれている？
            </div>
          </div>
          <!-- Typing Indicator inside history container for layout consistency -->
          <div class="typing-indicator" id="loading">
            <span></span>
            <span></span>
            <span></span>
          </div>
        </div>
        <div class="input-area">
          <input type="text" id="user-input" placeholder="AIに質問する..." onkeydown="if(event.key==='Enter') sendChat()">
          <button onclick="sendChat()" id="send-btn" title="送信">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
              <line x1="22" y1="2" x2="11" y2="13"></line>
              <polygon points="22 2 15 22 11 13 2 9 22 2"></polygon>
            </svg>
          </button>
        </div>
        <script>
          function sendChat() {
            const input = document.getElementById('user-input');
            const text = input.value.trim();
            if (!text) return;
            
            addMessage(text, 'user');
            input.value = '';
            document.getElementById('send-btn').disabled = true;
            
            // Move typing indicator to the bottom and show it
            const history = document.getElementById('chat-history');
            const loading = document.getElementById('loading');
            history.appendChild(loading);
            loading.style.display = 'flex';
            history.scrollTop = history.scrollHeight;

            google.script.run
              .withSuccessHandler(response => {
                loading.style.display = 'none';
                addMessage(response, 'ai');
                document.getElementById('send-btn').disabled = false;
              })
              .withFailureHandler(error => {
                loading.style.display = 'none';
                addMessage('エラーが発生しました: ' + error.message, 'ai');
                document.getElementById('send-btn').disabled = false;
              })
              .processChatQuery(text);
          }

          function parseMarkdown(text) {
            // Escape HTML to prevent XSS
            let safeText = text
              .replace(/&/g, "&amp;")
              .replace(/</g, "&lt;")
              .replace(/>/g, "&gt;")
              .replace(/"/g, "&quot;")
              .replace(/'/g, "&#039;");
            
            // Handle bold (**text**)
            safeText = safeText.replace(/\*\*(.*?)\*\*/g, "<strong>$1</strong>");
            
            // Handle lists (lines starting with - or * or numbers like 1.)
            const lines = safeText.split("\n");
            let inList = false;
            for (let i = 0; i < lines.length; i++) {
              const line = lines[i].trim();
              if (line.startsWith("- ") || line.startsWith("* ")) {
                const content = line.substring(2);
                if (!inList) {
                  lines[i] = "<ul><li>" + content + "</li>";
                  inList = true;
                } else {
                  lines[i] = "<li>" + content + "</li>";
                }
              } else {
                if (inList) {
                  lines[i] = "</ul>" + line;
                  inList = false;
                }
              }
            }
            if (inList) {
              lines.push("</ul>");
            }
            safeText = lines.join("\n");
            
            // Handle newlines
            safeText = safeText.replace(/\n/g, "<br>");
            
            // Cleanup consecutive empty lists
            safeText = safeText.replace(/<\/ul><br><ul>/g, "");
            
            return safeText;
          }

          function addMessage(text, type) {
            const history = document.getElementById('chat-history');
            const loading = document.getElementById('loading');
            
            const div = document.createElement('div');
            div.className = 'message ' + type;
            
            if (type === 'ai') {
              div.innerHTML = parseMarkdown(text);
            } else {
              // User message is plain text
              div.textContent = text;
            }
            
            // Insert before the loading indicator
            history.insertBefore(div, loading);
            history.scrollTop = history.scrollHeight;
          }
        </script>
      </body>
    </html>
  `;
  const html = HtmlService.createHtmlOutput(htmlContent).setTitle('🤖 AIデータ分析チャット').setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

// チャット処理バックエンド
function processChatQuery(userQuestion) {
  const currentSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const config = getConfig(currentSpreadsheet);
  if (!config) return "設定エラー：基礎データシートを確認してください。";

  const dataObj = getTargetData(config);
  if (!dataObj) return "データ取得エラー：対象のデータが見つかりません。";

  // チャット用プロンプトでAI呼び出し
  return askGeminiChat(dataObj.comments, config.apiKey, userQuestion, config.instruction);
}


// ==========================================
// 共通・ヘルパー関数群
// ==========================================

function getConfig(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(CONFIG_SHEET_NAME);
  if (!sheet) { return null; }
  
  const apiKeyCell = sheet.getRange('B1');
  let apiKeyValue = apiKeyCell.getValue().toString().trim();
  let apiKey = "";
  
  if (apiKeyValue === "1本格納しています。") {
    apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  } else if (apiKeyValue !== "") {
    apiKey = apiKeyValue;
    PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', apiKey);
    apiKeyCell.setValue("1本格納しています。");
    SpreadsheetApp.flush();
  } else {
    apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (apiKey) {
      apiKeyCell.setValue("1本格納しています。");
      SpreadsheetApp.flush();
    }
  }
  
  const url = sheet.getRange('B3').getValue();
  const sheetName = sheet.getRange('B4').getValue();
  const colLetter = sheet.getRange('B5').getValue();
  const instruction = sheet.getRange('B6').getValue(); 
  
  if (!apiKey || !url || !sheetName || !colLetter) return null;
  
  try {
    const targetSpreadsheetId = SpreadsheetApp.openByUrl(url).getId(); 
    return { apiKey, url, sheetName, colLetter, instruction, targetSpreadsheetId };
  } catch (e) { return null; }
}

function getTargetData(config) {
  const ss = SpreadsheetApp.openByUrl(config.url);
  const sheet = ss.getSheetByName(config.sheetName);
  if (!sheet) return null;
  const colStr = config.colLetter.toString().toUpperCase();
  let colIndex = 0;
  for (let i = 0; i < colStr.length; i++) {
    colIndex += (colStr.charCodeAt(i) - 64) * Math.pow(26, colStr.length - i - 1);
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const comments = sheet.getRange(2, colIndex, lastRow - 1, 1).getValues().flat().filter(t => t && t.toString().trim() !== "");
  if (comments.length === 0) return null;
  return { comments, fileName: ss.getName() };
}

// ==========================================
// Gemini API 関連
// ==========================================

// 1. 分析用（JSON出力）
function analyzeWithGemini(comments, apiKey, instruction) {
  const textData = comments.join("\n---\n");
  const instructionText = instruction ? `【重要：分析の視点】"${instruction}"` : "";
  const prompt = `
    あなたは教育データアナリストです。以下のアンケート自由記述を分析しJSONを出力してください。
    ${instructionText}
    【タスク】
    1. 全体の傾向を分析してください。
    2. 回答の内容から、主要な話題・テーマを「3つのカテゴリ」に自動分類してください。
    【出力JSON構造】
    {
      "summary": "全体の要約（200文字程度）",
      "sentiment_counts": { "positive": 数, "negative": 数, "neutral": 数 },
      "top_keywords": [ {"word": "単語1", "count": 数}, {"word": "単語2", "count": 数}, {"word": "単語3", "count": 数}, {"word": "単語4", "count": 数}, {"word": "単語5", "count": 数} ],
      "top_positive_points": ["良い点1", "良い点2", "良い点3"],
      "top_issues": ["課題1", "課題2", "課題3"],
      "auto_categories": [
        { "name": "カテゴリ名A", "summary": "要約", "issues": ["課題1", "課題2"] },
        { "name": "カテゴリ名B", "summary": "要約", "issues": ["課題1", "課題2"] },
        { "name": "カテゴリ名C", "summary": "要約", "issues": ["課題1", "課題2"] }
      ]
    }
    【データ】${textData}
  `;
  return callGemini(apiKey, prompt, true);
}

// 2. ドキュメント用（JSON出力）
function generateDocContentWithGemini(comments, apiKey, instruction) {
  const textData = comments.join("\n---\n");
  const context = instruction ? instruction : "学校運営に関するアンケート";
  const prompt = `
    あなたは、保護者に信頼されている誠実な学校管理者（校長または担当教諭）です。
    集まったアンケートの自由記述をもとに、保護者へ向けた「お便り」の原稿を作成し、JSON形式で出力してください。
    【背景・文脈】${context}
    【出力JSON構造】
    {
      "title": "文書のタイトル",
      "opening_text": "挨拶、全体の感謝、ポジティブな成果についての文章（数段落）",
      "qa_table": [
        { "category": "課題カテゴリ", "voice": "保護者からの声（要約）", "response": "対応方針" },
        { "category": "課題カテゴリ", "voice": "保護者からの声（要約）", "response": "対応方針" },
        { "category": "課題カテゴリ", "voice": "保護者からの声（要約）", "response": "対応方針" }
      ],
      "closing_text": "今後の決意、結びの言葉"
    }
    【要件】qa_tableには、特に重要と思われる課題を3〜4つ抽出。
    【アンケートデータ】${textData}
  `;
  return callGemini(apiKey, prompt, true); 
}

// 3. ★チャット用（テキスト出力）
function askGeminiChat(comments, apiKey, question, instruction) {
  const textData = comments.join("\n---\n");
  const context = instruction ? `(前提条件: ${instruction})` : "";
  
  const prompt = `
    あなたは教育データアナリストです。以下のアンケートデータ全体を読み込み、ユーザーからの質問に答えてください。
    
    【データ】
    ${textData}

    【ユーザーの質問】
    ${question} ${context}

    【回答のルール】
    - データに基づいた事実のみを回答してください。
    - データにないことは「データには含まれていません」と答えてください。
    - 簡潔で分かりやすい日本語で答えてください。
  `;
  return callGemini(apiKey, prompt, false);
}

// API呼び出し共通関数（安全対策・リトライ付き）
function callGemini(apiKey, prompt, isJson) {
  const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent?key=${apiKey}`;
  const payload = {
    "contents": [{ "parts": [{ "text": prompt }] }],
    "safetySettings": [
        {"category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_ONLY_HIGH"},
        {"category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_ONLY_HIGH"},
        {"category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_ONLY_HIGH"},
        {"category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_ONLY_HIGH"}
    ],
    "generationConfig": isJson ? { "response_mime_type": "application/json" } : {}
  };
  const options = {
    "method": "post", "contentType": "application/json",
    "payload": JSON.stringify(payload), "muteHttpExceptions": true
  };

  let maxRetries = 3;
  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    try {
      const response = UrlFetchApp.fetch(apiUrl, options);
      const statusCode = response.getResponseCode();
      const contentText = response.getContentText();
      const jsonResponse = JSON.parse(contentText);

      if (statusCode === 429 || (jsonResponse.error && jsonResponse.error.code === 429)) throw new Error("QuotaExceeded");
      if (jsonResponse.error) throw new Error(`Gemini API Error: ${jsonResponse.error.message}`);
      if (!jsonResponse.candidates || jsonResponse.candidates.length === 0) {
        const blockReason = jsonResponse.promptFeedback ? jsonResponse.promptFeedback.blockReason : "Unknown";
        throw new Error(`AIが応答を生成できませんでした（理由: ${blockReason}）。`);
      }

      const text = jsonResponse.candidates[0].content.parts[0].text;
      return isJson ? JSON.parse(text) : text;

    } catch (e) {
      if (e.message.includes("QuotaExceeded") || e.message.includes("429")) {
        if (attempt === maxRetries) throw new Error("アクセス集中により失敗しました。時間を空けて試してください。");
        const waitSeconds = 12 * (attempt + 1);
        Utilities.sleep(waitSeconds * 1000);
      } else {
        throw e;
      }
    }
  }
}

// ==========================================
// 出力生成 関連
// ==========================================

function createResponseDoc(data, titleSuffix, targetFolder, instruction) {
  const docName = `回答案_${titleSuffix}_${new Date().toLocaleDateString()}`;
  const doc = DocumentApp.create(docName);
  const body = doc.getBody();

  const titleText = data.title || `【下書き】${instruction || "アンケート"}への回答案`;
  const headerPara = body.insertParagraph(0, titleText);
  headerPara.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  headerPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendParagraph("\n");
  if (data.opening_text) body.appendParagraph(data.opening_text);

  if (data.qa_table && data.qa_table.length > 0) {
    body.appendParagraph("\n【主なご意見と学校の対応】");
    const tableData = [["課題・テーマ", "頂いたご意見（概要）", "学校の対応方針"]];
    data.qa_table.forEach(row => { tableData.push([row.category, row.voice, row.response]); });
    const table = body.appendTable(tableData);
    const headerRow = table.getRow(0);
    for (let i = 0; i < headerRow.getNumCells(); i++) {
      headerRow.getCell(i).setBackgroundColor('#eeeeee').getChild(0).asParagraph().setBold(true);
    }
    table.setColumnWidth(0, 100); table.setColumnWidth(1, 150); table.setColumnWidth(2, 200);
  }

  body.appendParagraph("\n");
  if (data.closing_text) body.appendParagraph(data.closing_text);

  body.appendHorizontalRule();
  const footerText = body.appendParagraph(`\n※この文章はAI（Gemini）によって生成された下書きです。\n学校の実情に合わせて、必ず加筆・修正を行ってから発行してください。\n元データ: ${titleSuffix}`);
  footerText.setForegroundColor("#666666").setFontSize(9);

  doc.saveAndClose();
  const file = DriveApp.getFileById(doc.getId());
  file.moveTo(targetFolder);
  return file.getUrl();
}

function createSentimentChartBlob(data) {
  try {
    const counts = data.sentiment_counts;
    const dataTable = Charts.newDataTable()
      .addColumn(Charts.ColumnType.STRING, "Category").addColumn(Charts.ColumnType.NUMBER, "Count")
      .addRow(["ポジティブ", counts.positive]).addRow(["ニュートラル", counts.neutral]).addRow(["ネガティブ", counts.negative])
      .build();
    return Charts.newPieChart().setDataTable(dataTable).setOption('title', '回答の傾向').setOption('colors', ['#34a853', '#d9d9d9', '#ea4335']).setOption('pieSliceTextStyle', {color: '#000000'}).setOption('is3D', true).setDimensions(400, 300).build().getBlob();
  } catch (e) { return null; }
}
function createKeywordChartBlob(data) {
  try {
    const keywords = data.top_keywords;
    if (!keywords || keywords.length === 0) return null;
    const dataTableBuilder = Charts.newDataTable().addColumn(Charts.ColumnType.STRING, "Keyword").addColumn(Charts.ColumnType.NUMBER, "Count");
    keywords.forEach(k => dataTableBuilder.addRow([k.word, k.count]));
    return Charts.newBarChart().setDataTable(dataTableBuilder.build()).setOption('title', '頻出・重要キーワード TOP5').setOption('legend', {position: 'none'}).setOption('colors', ['#4285f4']).setDimensions(400, 300).build().getBlob();
  } catch (e) { return null; }
}
function createSlidePresentation(data, titleSuffix, targetFolder, sentimentBlob, keywordBlob, instruction) {
  const fileName = `分析レポート_${titleSuffix}_${new Date().toLocaleDateString()}`;
  const presentation = SlidesApp.create(fileName);
  const titleSlide = presentation.getSlides()[0];
  titleSlide.getPageElements()[0].asShape().getText().setText(`アンケート自由記述 分析レポート`);
  let subtitle = `データ元: ${titleSuffix}\n作成日: ${new Date().toLocaleDateString()}`;
  if (instruction) subtitle += `\n\n※注力分析テーマ: ${instruction}`;
  titleSlide.getPageElements()[1].asShape().getText().setText(subtitle);
  const slide2 = presentation.appendSlide(SlidesApp.PredefinedLayout.TITLE_ONLY);
  slide2.getShapes()[0].getText().setText("全体概要とデータの傾向");
  const textArea = slide2.insertShape(SlidesApp.ShapeType.TEXT_BOX, 50, 60, 600, 100);
  textArea.getText().setText(`■ 全体の要約\n${data.summary}`);
  textArea.getText().getTextStyle().setFontSize(14);
  if (sentimentBlob) { const img1 = slide2.insertImage(sentimentBlob); img1.setLeft(20); img1.setTop(160); img1.setWidth(320); }
  if (keywordBlob) { const img2 = slide2.insertImage(keywordBlob); img2.setLeft(360); img2.setTop(160); img2.setWidth(320); }
  const slide3 = presentation.appendSlide(SlidesApp.PredefinedLayout.TITLE_AND_BODY);
  slide3.getShapes()[0].getText().setText("主な評価点（Good）");
  slide3.getShapes()[1].getText().setText(data.top_positive_points.map(p => `・${p}`).join("\n\n"));
  const slide4 = presentation.appendSlide(SlidesApp.PredefinedLayout.TITLE_AND_BODY);
  slide4.getShapes()[0].getText().setText("主な課題と改善のヒント");
  slide4.getShapes()[1].getText().setText(data.top_issues.map(p => `・${p}`).join("\n\n"));
  if (data.auto_categories && data.auto_categories.length > 0) {
    data.auto_categories.forEach(cat => {
      const catSlide = presentation.appendSlide(SlidesApp.PredefinedLayout.TITLE_AND_BODY);
      catSlide.getShapes()[0].getText().setText(`分野別分析: ${cat.name}`);
      const content = `■ 現状と要約\n${cat.summary}\n\n■ この分野の主な課題\n${cat.issues.map(i => `・${i}`).join('\n')}`;
      const bodyShape = catSlide.getShapes()[1];
      bodyShape.getText().setText(content);
      bodyShape.getText().getTextStyle().setFontSize(14);
    });
  }
  presentation.saveAndClose();
  const file = DriveApp.getFileById(presentation.getId());
  file.moveTo(targetFolder);
  return file.getUrl();
}
function saveResultToSheet(spreadsheet, data, fileName, sheetName, fileUrl, slideUrl, instruction) {
  let sheet = spreadsheet.getSheetByName(RESULT_SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(RESULT_SHEET_NAME);
    sheet.appendRow(['実行日時', '対象ファイル', '追加指示', '分野別分析概要', '要約', 'ポジティブ', 'ニュートラル', 'ネガティブ', '評価点', '課題点', 'スライドURL']);
    sheet.getRange(1, 1, 1, 11).setFontWeight('bold').setBackground('#d9ead3').setBorder(true, true, true, true, true, true);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 120); sheet.setColumnWidth(2, 200); sheet.setColumnWidth(3, 100);
    sheet.setColumnWidth(4, 250); sheet.setColumnWidth(5, 300); sheet.setColumnWidth(9, 300); sheet.setColumnWidth(10, 300);
  }
  const hyperLinkFormula = `=HYPERLINK("${fileUrl}", "${fileName} (${sheetName})")`;
  let catText = "";
  
  // ★ ここが修正された箇所です（「.substring(0, 50)...」を削除し、全文出力するように変更しました）
  if (data.auto_categories) { catText = data.auto_categories.map(c => `【${c.name}】${c.summary}`).join("\n"); }
  
  sheet.appendRow([ new Date(), hyperLinkFormula, instruction || "", catText, data.summary, data.sentiment_counts.positive, data.sentiment_counts.neutral, data.sentiment_counts.negative, data.top_positive_points.join('\n'), data.top_issues.join('\n'), slideUrl ]);
  const lastRow = sheet.getLastRow();
  const targetRange = sheet.getRange(lastRow, 1, 1, 11);
  targetRange.setVerticalAlignment('top').setWrap(true).setBorder(true, true, true, true, true, true);
  if (Number(data.sentiment_counts.negative) > Number(data.sentiment_counts.positive)) targetRange.setBackground('#ffe6e6');
}
function showCompleteModal(ui, url, typeName, sheetName) {
  const html = HtmlService.createHtmlOutput(`<p>完了しました！</p><ul><li>${typeName}が保存されました。</li><li>「${sheetName}」に記録されました。</li></ul><p><a href="${url}" target="_blank">作成された${typeName}を開く</a></p>`).setWidth(400).setHeight(250);
  ui.showModalDialog(html, '処理完了');
}