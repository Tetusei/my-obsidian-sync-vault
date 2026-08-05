/**
 * AppOps.gs
 * 公文書作成、議事録清書、名簿照合、カレンダー連携、マニュアル日誌、タスク整理
 */

/* --- フォルダ自動生成＆記憶の共通関数 --- */
function getOrCreateFolder(ss, masterSheet, cellPos, folderName) {
  let folderUrl = masterSheet.getRange(cellPos).getValue();
  let targetFolder;
  if (folderUrl && folderUrl.includes('drive.google.com/')) {
    try { targetFolder = DriveApp.getFolderById(folderUrl.match(/folders\/([-\w]+)/)[1]); } catch(e) {}
  }
  if (!targetFolder) {
    const parentFolder = DriveApp.getFileById(ss.getId()).getParents().next();
    const folders = parentFolder.getFoldersByName(folderName);
    targetFolder = folders.hasNext() ? folders.next() : parentFolder.createFolder(folderName);
    masterSheet.getRange(cellPos).setValue(targetFolder.getUrl()); 
  }
  return targetFolder;
}

/* --- 文書作成 --- */
function createSchoolDocument() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const docSheet = ss.getSheetByName(Config.SHEET_NAME_DOC_FACTORY);
  const masterSheet = ss.getSheetByName(Config.SHEET_NAME_MASTER);
  
  if (!docSheet) { SpreadsheetApp.getUi().alert("文書作成シートが見つかりません。"); return; }
  
  const title = docSheet.getRange(Config.DOC_FACTORY_POS.TITLE).getValue();
  const target = docSheet.getRange(Config.DOC_FACTORY_POS.TARGET).getValue();
  const content = docSheet.getRange(Config.DOC_FACTORY_POS.CONTENT).getValue();
  const roleLabel = docSheet.getRange("A8").getValue().toString().trim() || "代 表 者";
  
  if (!title || !content) { SpreadsheetApp.getUi().alert("「タイトル」と「内容」を入力してください。"); return; }
  ss.toast("文書の体裁を整えています...", "🪄 文書構成中");

  const apiKey = getApiKey();
  const modelName = masterSheet.getRange(Config.MASTER_POS.MODEL_NAME_CELL).getValue() || "gemini-1.5-flash";
  
  const prompt = `あなたは日本の公立学校の管理職をサポートする文書作成アシスタントです。
【文書のタイトル】: ${title}\n【宛先】: ${target}\n【盛り込みたい内容】: ${content}
【文体・トーン】: ${docSheet.getRange(Config.DOC_FACTORY_POS.TONE).getValue()}
【発行元名称】: ${docSheet.getRange(Config.DOC_FACTORY_POS.SCHOOL_NAME).getValue()}
【責任者氏名】: ${docSheet.getRange(Config.DOC_FACTORY_POS.PRINCIPAL_NAME).getValue()}
【ルール】時候の挨拶を含め正式な公文書に。役職名はシートから取得した「${roleLabel}」を使用すること。出力時は「${roleLabel} ◯◯ ◯◯」のように役職名と氏名の間に全角スペースを入れ、役職名自体の文字間も公文書として美しい体裁（例：代 表 者、校 長 など）に整えること。解説不要で本文のみ。`;

  const draftText = callGeminiAPI(prompt, apiKey, modelName);
  if (!draftText) { SpreadsheetApp.getUi().alert("通信に失敗しました。"); return; }

  const doc = DocumentApp.create(title);
  doc.getBody().setText(draftText);
  
  const targetFolder = getOrCreateFolder(ss, masterSheet, Config.MASTER_POS.DRAFT_FOLDER_URL_CELL, Config.DRAFT_FOLDER_NAME);
  const docFile = DriveApp.getFileById(doc.getId());
  docFile.moveTo(targetFolder);
  
  docSheet.getRange(Config.DOC_FACTORY_POS.RESULT_URL).setValue(docFile.getUrl());
  ss.toast("汎用性の高いドラフトが完成しました！", "✅ 作成完了");
}

/* --- 議事録清書 --- */
function finalizeMinutes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(Config.SHEET_NAME_MASTER);
  const ui = SpreadsheetApp.getUi();

  const folderUrl = masterSheet.getRange(Config.MASTER_POS.MINUTES_MEMO_FOLDER_URL_CELL).getValue();
  if (!folderUrl) { ui.alert("B13セルに「会議メモ」フォルダのURLを入力してください。"); return; }
  ss.toast("フォルダ内の最新ファイルを探しています...", "🔍 検索中");

  try {
    const memoFolder = DriveApp.getFolderById(folderUrl.split('/folders/')[1].split('?')[0]);
    const files = memoFolder.getFiles();
    let latestFile = null, lastUpdated = 0;

    while (files.hasNext()) {
      const file = files.next();
      const mType = file.getMimeType();
      if (mType === MimeType.GOOGLE_DOCS || mType === MimeType.PDF) {
        if (file.getLastUpdated().getTime() > lastUpdated) { lastUpdated = file.getLastUpdated().getTime(); latestFile = file; }
      }
    }
    if (!latestFile) { ui.alert("ドキュメントまたはPDFが見つかりませんでした。"); return; }

    let promptText = "", fileData = null;
    if (latestFile.getMimeType() === MimeType.GOOGLE_DOCS) {
      const rawText = DocumentApp.openById(latestFile.getId()).getBody().getText();
      promptText = `以下の会議メモを清書してください：\n\n${rawText}`;
    } else {
      fileData = { mime_type: "application/pdf", data: Utilities.base64Encode(latestFile.getBlob().getBytes()) };
      promptText = "添付されたPDFファイルの内容を読み取り、議事録として清書してください。";
    }

    ss.toast(`「${latestFile.getName()}」を分析中...`, "🪄 AI秘書");
    const apiKey = getApiKey();
    const modelName = masterSheet.getRange(Config.MASTER_POS.MODEL_NAME_CELL).getValue() || "gemini-1.5-flash";
    const finalPrompt = `あなたは学校組織の事務局担当です。提供情報を元に正式な「議事録」を作成してください。
【ルール】「会議概要」「決定事項」「今後の課題」「担当タスク」で整理。簡潔な公的文体に。余計な解説は不要。`;

    const formattedText = callGeminiAPI_v2(finalPrompt + "\n\n" + promptText, apiKey, modelName, fileData);
    if (!formattedText) { ui.alert("通信失敗。ファイルサイズ超過等の可能性があります。"); return; }

    const finalDoc = DocumentApp.create(`【清書版】${latestFile.getName().replace(".pdf", "")}`);
    finalDoc.getBody().setText(formattedText);

    const minutesFolder = getOrCreateFolder(ss, masterSheet, Config.MASTER_POS.MINUTES_SAVE_FOLDER_URL_CELL, Config.MINUTES_FOLDER_NAME);
    DriveApp.getFileById(finalDoc.getId()).moveTo(minutesFolder);
    
    const htmlOutput = HtmlService.createHtmlOutput(
      `<div style="text-align:center; padding:10px;">
         <p style="font-size: 14px; margin-bottom: 20px;">対象: ${latestFile.getName()}</p>
         <a href="${finalDoc.getUrl()}" target="_blank" style="padding:10px 20px; background:#4285f4; color:white; text-decoration:none; border-radius:5px; font-weight:bold;">📝 議事録を開く</a>
         <p style="font-size: 11px; color: gray; margin-top: 20px;">保存先フォルダ: ${Config.MINUTES_FOLDER_NAME}</p>
       </div>`
    ).setWidth(350).setHeight(180);
    ui.showModelessDialog(htmlOutput, "✅ 清書完了");

  } catch (e) { ui.alert("エラー：" + e.toString()); }
}

/* --- 名簿照合 --- */
function runRosterCheck() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(Config.SHEET_NAME_MASTER);
  
  const rosterUrl = masterSheet.getRange(Config.MASTER_POS.ROSTER_FILE_URL_CELL).getValue();
  const folderUrl = masterSheet.getRange(Config.MASTER_POS.WORK_FOLDER_URL_CELL).getValue();
  const apiKey = getApiKey();
  const modelName = masterSheet.getRange(Config.MASTER_POS.MODEL_NAME_CELL).getValue();

  if (!rosterUrl || !folderUrl) { SpreadsheetApp.getUi().alert("B14, B15にURLを入力してください。"); return; }
  ss.toast("データ取得中...", "🔍 準備中");

  const rosterData = getRosterData(rosterUrl);
  const folder = DriveApp.getFolderById(folderUrl.match(/folders\/([-\w]+)/)[1]);
  const files = folder.getFiles();
  let checkResults = [];

  while (files.hasNext()) {
    const file = files.next();
    if (file.getName().includes("照合結果レポート")) continue;
    ss.toast(`${file.getName()} 解析中...`);
    try {
      const fileText = extractTextFromFile(file);
      if (fileText) checkResults.push(askGeminiForVerification(fileText, rosterData, file.getName(), apiKey, modelName));
    } catch (e) { checkResults.push(`### 【エラー】${file.getName()}\n${e.toString()}`); }
  }

  if (checkResults.length > 0) {
    const doc = DocumentApp.create(`【AI照合結果レポート】_${Utilities.formatDate(new Date(), "JST", "yyyyMMdd_HHmm")}`);
    const body = doc.getBody();
    body.appendParagraph(doc.getName()).setHeading(DocumentApp.ParagraphHeading.HEADING1);
    checkResults.forEach(res => { body.appendParagraph(res); body.appendHorizontalRule(); });
    DriveApp.getFileById(doc.getId()).moveTo(folder);
    SpreadsheetApp.getUi().showModelessDialog(HtmlService.createHtmlOutput(`<p>照合完了！</p><a href="${doc.getUrl()}" target="_blank">結果を開く</a>`), "✅ 完了");
  } else { ss.toast("対象ファイルなし", "⚠️ 中断"); }
}

/* --- マニュアル日誌作成 --- */
function createDailyReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(Config.SHEET_NAME_DASHBOARD);
  const dateValue = sheet.getRange("B1").getValue();
  if (!(dateValue instanceof Date)) { SpreadsheetApp.getUi().alert("B1セルに日付を入力してください。"); return; }
  
  const scheduleData = sheet.getRange(4, 2, 17, 4).getValues();
  const taskText = getCompletedTasksToday(ss); 
  const generalMemos = sheet.getRange(22, 2, 20, 1).getValues().filter(r => r[0] !== "" && r[0] !== null);

  const docTitle = `【教頭日誌】${Utilities.formatDate(dateValue, "JST", "yyyy年MM月dd日")}`;
  const doc = DocumentApp.create(docTitle);
  const body = doc.getBody();

  body.appendParagraph(docTitle).setHeading(DocumentApp.ParagraphHeading.HEADING1).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  
  body.appendParagraph("１．主な動き（予定）").setHeading(DocumentApp.ParagraphHeading.HEADING2);
  let hasSchedule = false;
  scheduleData.forEach(row => {
    if (row[2] && row[2].toString().trim() !== "") {
      const s = row[0], e = row[1], memo = row[3];
      const sStr = (s instanceof Date) ? Utilities.formatDate(s, "JST", "HH:mm") : (s || "").toString();
      const eStr = (e instanceof Date) ? Utilities.formatDate(e, "JST", "HH:mm") : (e || "").toString();
      let tLabel = (sStr === "終日" || sStr === "" || (sStr === "0:00" && (eStr === "" || eStr === "0:00"))) ? "【終日】" : `【${sStr}〜${eStr}】`;
      body.appendListItem(`${tLabel} ${row[2]}${memo ? ' （内容：' + memo + '）' : ''}`).setGlyphType(DocumentApp.GlyphType.BULLET);
      hasSchedule = true;
    }
  });
  if (!hasSchedule) body.appendParagraph("（特になし）");

  body.appendParagraph("２．業務実績（タスク）").setHeading(DocumentApp.ParagraphHeading.HEADING2);
  if (taskText && taskText !== "本日完了した業務はありません。") {
    taskText.split('\n').forEach(t => body.appendListItem(t.replace('・', '')).setGlyphType(DocumentApp.GlyphType.BULLET));
  } else {
    body.appendParagraph("（特になし）");
  }

  body.appendParagraph("３．所感").setHeading(DocumentApp.ParagraphHeading.HEADING2);
  if (generalMemos.length > 0) {
    generalMemos.forEach(m => body.appendListItem(m[0]).setGlyphType(DocumentApp.GlyphType.SQUARE_BULLET));
  } else {
    body.appendParagraph("（特になし）");
  }

  body.appendHorizontalRule();
  body.appendParagraph(`報告者：教頭\n作成日：${Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm")}`).setAlignment(DocumentApp.HorizontalAlignment.RIGHT);

  const masterSheet = ss.getSheetByName(Config.SHEET_NAME_MASTER);
  const targetFolder = getOrCreateFolder(ss, masterSheet, Config.MASTER_POS.REPORT_FOLDER_URL_CELL, Config.REPORT_FOLDER_NAME);
  DriveApp.getFileById(doc.getId()).moveTo(targetFolder);

  const htmlOutput = HtmlService.createHtmlOutput(
    `<div style="text-align:center; padding:10px;"><p>マニュアル日誌が完成しました。</p><a href="${doc.getUrl()}" target="_blank" style="padding:10px 20px; background:#4285f4; color:white; text-decoration:none; border-radius:5px;">日誌を開く</a></div>`
  ).setWidth(350).setHeight(150);
  SpreadsheetApp.getUi().showModelessDialog(htmlOutput, "✅ 作成完了");
}

/* --- ToDoカレンダー連携機能 --- */
function addSingleTaskToCalendar(title, pic, dueDateStr, content) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(Config.SHEET_NAME_MASTER);
  const todoCalId = masterSheet.getRange(Config.MASTER_POS.TODO_CALENDAR_ID_CELL).getValue();
  
  if (!todoCalId || !title) return;

  try {
    const calendar = CalendarApp.getCalendarById(todoCalId);
    if (!calendar) return;

    let dueDate;
    if (dueDateStr instanceof Date) {
      dueDate = dueDateStr;
    } else if (typeof dueDateStr === 'string' && dueDateStr.trim() !== '') {
      dueDate = new Date(dueDateStr);
    } else {
      return; 
    }

    if (isNaN(dueDate.getTime())) return;

    const eventTitle = `📌[ToDo] ${title} (${pic || "未設定"})`;
    if (calendar.getEventsForDay(dueDate, {search: eventTitle}).length === 0) {
      calendar.createAllDayEvent(eventTitle, dueDate, {description: content || ""});
    }
  } catch(e) {
    console.log("カレンダー登録エラー: " + e.message);
  }
}

function syncTasksToCalendar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const calendarId = ss.getSheetByName(Config.SHEET_NAME_MASTER).getRange(Config.MASTER_POS.TODO_CALENDAR_ID_CELL).getValue();
  if (!calendarId) {
    SpreadsheetApp.getUi().alert("基礎データのB18セルに「ToDo専用カレンダー」のIDを入力してください。");
    return;
  }
  
  const calendar = CalendarApp.getCalendarById(calendarId);
  const data = ss.getSheetByName(Config.SHEET_NAME_TODO).getDataRange().getValues();
  data.shift(); 
  data.forEach(row => {
    const title = row[Config.TODO_COL.TITLE], dueDate = row[Config.TODO_COL.DUE_DATE];
    if (dueDate instanceof Date && row[Config.TODO_COL.STATUS] !== '完了' && title) {
      const eventTitle = `📌[ToDo] ${title} (${row[Config.TODO_COL.PIC] || "未設定"})`;
      if (calendar.getEventsForDay(dueDate, {search: eventTitle}).length === 0) {
        calendar.createAllDayEvent(eventTitle, dueDate, {description: row[Config.TODO_COL.CONTENT] || ""});
      }
    }
  });
}

function getTodayEventsText() {
  const masterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(Config.SHEET_NAME_MASTER);
  const calendarId = masterSheet.getRange(Config.MASTER_POS.CALENDAR_ID_CELL).getValue();
  if (!calendarId) return "⚠️カレンダーID未設定";
  try {
    const events = CalendarApp.getCalendarById(calendarId).getEvents(new Date(), new Date(new Date().setHours(23, 59, 59, 999)));
    if (events.length === 0) return "・本日の予定はありません。";
    return events.map(e => `・[${e.isAllDayEvent() ? "終日" : Utilities.formatDate(e.getStartTime(), "JST", "HH:mm")}] ${e.getTitle()}`).join('\n');
  } catch (e) { return "⚠️取得失敗"; }
}

/* --- Gemini API 共通関数群 --- */
function callGeminiAPI(prompt, apiKey, modelName) {
  const payload = { "contents": [{ "parts": [{ "text": prompt }] }] };
  try {
    const res = callGeminiWithRotation(payload, modelName);
    return JSON.parse(res.getContentText()).candidates[0].content.parts[0].text;
  } catch (e) {
    console.error("callGeminiAPI error:", e);
    return null;
  }
}

function callGeminiAPI_v2(prompt, apiKey, modelName, fileData = null) {
  const parts = [{ "text": prompt }];
  if (fileData) parts.push({ "inline_data": fileData });
  const payload = { "contents": [{ "parts": parts }] };
  try {
    const res = callGeminiWithRotation(payload, modelName);
    return JSON.parse(res.getContentText()).candidates[0].content.parts[0].text;
  } catch (e) {
    console.error("callGeminiAPI_v2 error:", e);
    return null;
  }
}

/* --- 名簿ヘルパー --- */
function getRosterData(url) {
  let combinedData = "";
  SpreadsheetApp.openByUrl(url).getSheets().forEach(sheet => {
    const values = sheet.getDataRange().getValues();
    if (values.length > 0) combinedData += `\n--- シート名: ${sheet.getName()} ---\n` + values.map(row => row.join(" ")).join("\n") + "\n";
  });
  return combinedData;
}
function extractTextFromFile(file) {
  const mimeType = file.getMimeType(), id = file.getId();
  if (mimeType === MimeType.GOOGLE_DOCS) return DocumentApp.openById(id).getBody().getText();
  if (mimeType === MimeType.PLAIN_TEXT) return file.getBlob().getDataAsString();
  try {
    const tempFile = Drive.Files.create({ name: "temp_ocr", mimeType: MimeType.GOOGLE_DOCS }, file.getBlob());
    const text = DocumentApp.openById(tempFile.id).getBody().getText();
    DriveApp.getFileById(tempFile.id).setTrashed(true);
    return text;
  } catch (e) { return null; }
}
function askGeminiForVerification(content, roster, fileName, apiKey, modelName) {
  const prompt = `あなたは校閲担当。名簿データ(正解)とファイル内容を照合し、漢字や出席番号の不一致を指摘して下さい。\n【名簿】\n${roster}\n【ファイル名】${fileName}\n【内容】\n${content}`;
  const res = callGeminiAPI(prompt, apiKey, modelName);
  return `## ファイル名: ${fileName}\n\n${res || "エラーが発生しました。"}`;
}

/* --- 📱 音声メモ（フォーム）自動処理 --- */
function processVoiceMemo(e) {
  if (!e || !e.values) return; 
  const rawText = e.values[1]; 
  if (!rawText) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(Config.SHEET_NAME_MASTER);
  const todoSheet = ss.getSheetByName(Config.SHEET_NAME_TODO);
  
  const apiKey = getApiKey();
  const modelName = masterSheet.getRange(Config.MASTER_POS.MODEL_NAME_CELL).getValue() || "gemini-1.5-flash";

  const prompt = `あなたは優秀な学校事務の秘書です。以下の教頭先生の音声入力メモを分析し、ToDoリスト用のデータに整理してください。
【音声メモ】: ${rawText}

【ルール】
以下の要素を「|」区切りで出力してください。解説や挨拶は一切不要です。
件名|内容|期限(yyyy/MM/dd形式、不明なら空白)|重要度(高/中/低、デフォルトは中)

【超重要ルール】
「内容」の文章中には、人間が見て一目でスケジュールがわかるよう、「〇月〇日(〇)までに」「明日の午前中に」といった【期限や日時の情報】を必ず自然な形で含めてください。`;

  let aiResponse = callGeminiAPI(prompt, apiKey, modelName);
  if (!aiResponse) aiResponse = `AI解析エラー|${rawText}||中`;
  
  const data = aiResponse.split('|');
  const subject = data[0] ? data[0].trim() : "（件名なし）";
  const content = data[1] ? data[1].trim() : rawText;
  const deadline = data[2] ? data[2].trim() : "";
  const priority = data[3] ? data[3].trim() : "中";
  const owner = "教頭"; 
  
  const today = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd");
  
  todoSheet.appendRow([
    today, "📱音声メモ", subject, content, owner, deadline, priority, "未着手", 
    false, "", "", "", `【元テキスト】\n${rawText}`
  ]);
  const lastRow = todoSheet.getLastRow();
  todoSheet.getRange(lastRow, Config.TODO_COL.ACTION + 1).insertCheckboxes();
  setTodoRowValidations(todoSheet, lastRow);

  if (deadline) {
    addSingleTaskToCalendar(subject, owner, deadline, content);
  }
}

function turnOnVoiceMemoTrigger() {
  const functionName = "processVoiceMemo";
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === functionName) ScriptApp.deleteTrigger(triggers[i]);
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger(functionName).forSpreadsheet(ss).onFormSubmit().create();
  SpreadsheetApp.getUi().alert("✅ 音声メモ連携を【ON】にしました！\n\nスマホのフォームから送信すると、自動でToDoシートへ追加されます。");
}

function turnOffVoiceMemoTrigger() {
  const functionName = "processVoiceMemo";
  const triggers = ScriptApp.getProjectTriggers();
  let found = false;
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === functionName) { ScriptApp.deleteTrigger(triggers[i]); found = true; }
  }
  if (found) SpreadsheetApp.getUi().alert("❌ 音声メモ連携を【OFF】にしました。\n\nフォームからの自動取り込みを停止しました。");
  else SpreadsheetApp.getUi().alert("ℹ️ 音声メモ連携はすでに【OFF】（未設定）になっています。");
}

/* --- ToDoシートの表示・非表示・削除切り替え --- */
function hideCompletedTasks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(Config.SHEET_NAME_TODO);
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  let hiddenCount = 0;
  
  for (let i = 1; i < data.length; i++) {
    const status = data[i][Config.TODO_COL.STATUS]; 
    // ✨【変更】「完了」という文字が含まれていれば非表示にする（完了(転送)にも対応）
    if (String(status).includes('完了')) { 
      sheet.hideRows(i + 1); 
      hiddenCount++; 
    } 
    else { 
      sheet.showRows(i + 1); 
    }
  }
  
  if (hiddenCount > 0) ss.toast(`完了したタスク ${hiddenCount} 件を非表示にしました！`, '🙈 整理完了', 3000);
  else ss.toast('隠すタスク（完了状態のもの）はありませんでした。', 'ℹ️ お知らせ', 3000);
}

function showAllTasks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(Config.SHEET_NAME_TODO);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) { sheet.showRows(2, lastRow - 1); }
  ss.toast('すべてのタスクを表示しました。', '👀 全表示', 3000);
}

// ✨ 【新機能】完了したToDoタスクを完全に削除する
function deleteCompletedTasks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  // 1. 確認メッセージを表示
  const response = ui.alert(
    '🗑️ 完了したToDoの完全削除',
    '状態が「完了」になっているタスクを完全に削除します。\n（※万が一のために、実行前に自動でバックアップを作成します）\n\nよろしいですか？',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    ss.toast('削除をキャンセルしました。', 'ℹ️ キャンセル', 3000);
    return;
  }

  // 2. 自動バックアップの実行
  ss.toast('バックアップを作成しています...', '📦 安全確保', 3000);
  createBackupCore(ss);
  
  // 3. 削除処理
  ss.toast('バックアップが完了しました。削除を実行します...', '🗑️ 処理中', 3000);
  const sheet = ss.getSheetByName(Config.SHEET_NAME_TODO);
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  let deletedCount = 0;

  // 下の行から順番に削除する（行ズレを防ぐため）
  for (let i = data.length - 1; i >= 1; i--) {
    const status = data[i][Config.TODO_COL.STATUS];
    // ✨【変更】「完了」という文字が含まれていれば削除対象にする
    if (String(status).includes('完了')) {
      sheet.deleteRow(i + 1);
      deletedCount++;
    }
  }

  // 4. 完了通知
  if (deletedCount > 0) {
    ss.toast(`完了したタスク ${deletedCount} 件を削除して整理しました！`, '🗑️ 完了', 4000);
  } else {
    ss.toast('削除するタスク（完了状態のもの）はありませんでした。', 'ℹ️ お知らせ', 3000);
  }
}



/**
 * ✨ 基礎データシートに記載されている各種フォルダを自動生成（存在すれば再利用）し、
 * そのURLを該当するセルに書き込む
 */
function ensureFoldersSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(Config.SHEET_NAME_MASTER);
  if (!masterSheet) return;
  
  // 1. 日報用フォルダ (B8)
  getOrCreateFolder(ss, masterSheet, Config.MASTER_POS.REPORT_FOLDER_URL_CELL, Config.REPORT_FOLDER_NAME);
  
  // 2. 会議メモ用フォルダ (B13)
  getOrCreateFolder(ss, masterSheet, Config.MASTER_POS.MINUTES_MEMO_FOLDER_URL_CELL, '会議メモ');
  
  // 3. 照合ワークスペース用フォルダ (B15)
  getOrCreateFolder(ss, masterSheet, Config.MASTER_POS.WORK_FOLDER_URL_CELL, '照合ワークスペース');
  
  // 4. 文書作成用フォルダ (B16)
  getOrCreateFolder(ss, masterSheet, Config.MASTER_POS.DRAFT_FOLDER_URL_CELL, Config.DRAFT_FOLDER_NAME);
  
  // 5. 議事録保存用フォルダ (B17)
  getOrCreateFolder(ss, masterSheet, Config.MASTER_POS.MINUTES_SAVE_FOLDER_URL_CELL, Config.MINUTES_FOLDER_NAME);
  
  // 6. メール添付ファイルフォルダ (B19)
  getOrCreateFolder(ss, masterSheet, Config.MASTER_POS.ATTACHMENTS_FOLDER_URL_CELL, 'メール添付ファイル');
}