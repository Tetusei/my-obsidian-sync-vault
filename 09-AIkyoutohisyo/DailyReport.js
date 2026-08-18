/**
 * DailyReport.gs
 * AI日報作成、タスク催促、チャット通知/受信
 */

function generateDailyReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(Config.SHEET_NAME_MASTER);
  if (!String(masterSheet.getRange(Config.MASTER_POS.REPORT_ON_OFF_CELL).getValue()).match(/^(ON|on|オン)$/i)) return;

  let targetFolder = getOrCreateFolder(ss, masterSheet, Config.MASTER_POS.REPORT_FOLDER_URL_CELL, Config.REPORT_FOLDER_NAME);
  const dashData = getDashboardData(ss);
  const taskText = getCompletedTasksToday(ss);

  const now = new Date();
  const dayOfWeek = ["日", "月", "火", "水", "木", "金", "土"][now.getDay()];
  const fiscalYearMatch = Config.FISCAL_YEAR.match(/令和(\d+)年/);
  const jpnYearStr = fiscalYearMatch ? `令和${fiscalYearMatch[1]}年` : `令和${now.getFullYear() - 2018}年`;
  const fullDateHeader = `${jpnYearStr}${Utilities.formatDate(now, "JST", "MM月dd日")}(${dayOfWeek})`;

  const modelName = masterSheet.getRange(Config.MASTER_POS.MODEL_NAME_CELL).getValue() || "gemini-1.5-flash";

  const reportContent = askGeminiForReport(dashData.schedule, taskText, dashData.memo, modelName, fullDateHeader);
  if (!reportContent) { ss.toast("AIでの日報生成に失敗しました。", "エラー"); return; }

  const docTitle = `【教頭日誌】${Utilities.formatDate(now, "JST", "yyyy年MM月dd日")}`;
  const doc = DocumentApp.create(docTitle);
  const body = doc.getBody();
  body.insertParagraph(0, docTitle).setAlignment(DocumentApp.HorizontalAlignment.CENTER).setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph(""); 
  body.appendParagraph(reportContent);
  
  const docFile = DriveApp.getFileById(doc.getId());
  docFile.moveTo(targetFolder);
  sendChatNotification(masterSheet, docTitle, docFile.getUrl());
  ss.toast("一般メモを反映してAI日報を作成しました！", "📝作成完了");
}

function sendTaskReminders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(Config.SHEET_NAME_MASTER);
  
  const isOn = masterSheet.getRange(Config.MASTER_POS.REMINDER_ON_OFF_CELL).getValue();
  if (!String(isOn).match(/^(ON|on|オン)$/i)) return;

  const now = new Date();
  const dayOfWeek = now.getDay();
  if (dayOfWeek === 0 || dayOfWeek === 6) return;
  try {
    const holidays = CalendarApp.getCalendarById("ja.japanese#holiday@group.v.calendar.google.com").getEventsForDay(now);
    if (holidays.length > 0) return;
  } catch(e) {}

  const webhookUrl = masterSheet.getRange(Config.MASTER_POS.ALL_WEBHOOK_CELL).getValue();
  if (!webhookUrl || !webhookUrl.startsWith('http')) return;

  const data = ss.getSheetByName(Config.SHEET_NAME_TODO).getDataRange().getValues();
  data.shift(); 
  const tomorrow = new Date(); tomorrow.setDate(tomorrow.getDate() + 1); tomorrow.setHours(0, 0, 0, 0);

  let reminderTasks = [];
  data.forEach(row => {
    const title = row[Config.TODO_COL.TITLE], pic = row[Config.TODO_COL.PIC], dueDateRaw = row[Config.TODO_COL.DUE_DATE];
    if (pic && !pic.includes('教頭') && !String(row[Config.TODO_COL.STATUS]).includes('完了') && dueDateRaw instanceof Date) {
      const dueDate = new Date(dueDateRaw.getTime()); dueDate.setHours(0, 0, 0, 0);
      if (dueDate.getTime() === tomorrow.getTime()) reminderTasks.push(`・**${title}** （担当: ${pic} 先生）`);
    }
  });

  if (reminderTasks.length > 0) {
    let messageText = `🔔 **【AI秘書からのリマインド】**\n\n先生方、毎日お疲れ様です。教頭先生のAI秘書です。\n明日が提出期限となっているタスクが ${reminderTasks.length}件 ございます。\nお忙しいところ大変恐縮ですが、ご確認とご対応をよろしくお願いいたします🙇‍♂️\n\n${reminderTasks.join('\n')}\n\n無事に完了されましたら、教頭先生までご一報ください。`;
    UrlFetchApp.fetch(webhookUrl, { "method": "post", "contentType": "application/json", "payload": JSON.stringify({ "text": messageText }) });
  }
}

function sendDailyMorningDigest() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(Config.SHEET_NAME_MASTER);
  
  const now = new Date();
  if (now.getDay() === 0 || now.getDay() === 6) return;
  try {
    const holidays = CalendarApp.getCalendarById("ja.japanese#holiday@group.v.calendar.google.com").getEventsForDay(now);
    if (holidays.length > 0) return;
  } catch(e) {}

  const webhookUrl = masterSheet.getRange(Config.MASTER_POS.WEBHOOK_CELL).getValue();
  if (!webhookUrl || !webhookUrl.startsWith('http')) return;

  const todaySchedule = getTodayEventsText();
  const data = ss.getSheetByName(Config.SHEET_NAME_TODO).getDataRange().getValues();
  const headers = data.shift();
  
  // 各列のインデックスを見つけるヘルパー関数（表記揺れや空白に対応、大文字小文字無視、部分一致）
  const findHeaderIndex = (keywords, defaultIdx) => {
    let idx = headers.findIndex(h => {
      const s = String(h).trim();
      return keywords.some(k => s === k);
    });
    if (idx !== -1) return idx;
    idx = headers.findIndex(h => {
      const s = String(h).trim();
      return keywords.some(k => s.includes(k));
    });
    return idx !== -1 ? idx : defaultIdx;
  };

  const colState = findHeaderIndex(['ステータス', '状態', '状況', '進捗'], Config.TODO_COL.STATUS);
  const colLimit = findHeaderIndex(['期限', '期日'], Config.TODO_COL.DUE_DATE);
  const colTitle = findHeaderIndex(['件名', 'タイトル', 'ToDo'], Config.TODO_COL.TITLE);
  const colPic = findHeaderIndex(['担当'], Config.TODO_COL.PIC);
  const colPriority = findHeaderIndex(['重要度', '優先'], Config.TODO_COL.PRIORITY);

  let urgentTasks = [], highTasks = [], normalTasks = [];
  const today = new Date(); today.setHours(0, 0, 0, 0);
  
  // 30日後の日付を計算（1ヶ月以上先のToDoを通知から除外するため）
  const thirtyDaysLater = new Date(today.getTime());
  thirtyDaysLater.setDate(thirtyDaysLater.getDate() + 30);

  data.forEach(row => {
    const statusVal = String(row[colState] || '');
    // ステータスに「完了」が含まれているもの、または件名が空のものは通知から除外
    if (statusVal.includes('完了') || !row[colTitle]) return;
    let limitDate = row[colLimit] instanceof Date ? row[colLimit] : null;
    
    // 期限が設定されており、それが30日（約1ヶ月）以上先の場合は通知をスキップする
    if (limitDate && limitDate.getTime() > thirtyDaysLater.getTime()) return;

    const taskObj = {
      title: row[colTitle], pic: row[colPic] || '未設定',
      limitStr: limitDate ? Utilities.formatDate(limitDate, "JST", "MM/dd") : "未設定",
      priority: row[colPriority]
    };
    if (limitDate && limitDate.getTime() <= today.getTime()) {
      taskObj.label = limitDate.getTime() < today.getTime() ? '🔴超過' : '🔥本日';
      urgentTasks.push(taskObj);
    } else if (taskObj.priority === '高') highTasks.push(taskObj);
    else normalTasks.push(taskObj);
  });

  let msg = `🌅 おはようございます！教頭先生のAI秘書です。\n${Config.FISCAL_YEAR} ${Utilities.formatDate(now, "JST", "M月d日(E)")}\n\n📅 **【本日の予定】**\n${todaySchedule}\n\n────────────────\n\n`;
  const formatList = (tasks) => tasks.map(t => `・${t.label ? `[${t.label}] ` : ""}${t.title} (担当: ${t.pic}, 期限: ${t.limitStr})`).join('\n');

  if (urgentTasks.length > 0) msg += `⚠️ **【至急・本日期限のToDo】**\n${formatList(urgentTasks)}\n\n`;
  if (highTasks.length > 0) msg += `⭐ **【重要ToDo（期限順）】**\n${formatList(highTasks)}\n\n`;
  if (normalTasks.length > 0 && (urgentTasks.length + highTasks.length < 5)) msg += `🏃 **【その他のToDo】**\n${formatList(normalTasks.slice(0, 5))}\n\n`;
  msg += `今日も一日よろしくお願いします！`;

  UrlFetchApp.fetch(webhookUrl, { "method": "post", "contentType": "application/json", "payload": JSON.stringify({ "text": msg }) });
}

function doPost(e) {
  try {
    const props = PropertiesService.getScriptProperties();
    const expectedToken = props.getProperty(Config.CHAT_WEBHOOK_TOKEN_PROPERTY);
    const receivedToken = e && e.parameter ? String(e.parameter.token || '') : '';

    if (!expectedToken || !receivedToken || receivedToken !== expectedToken) {
      console.warn('認証されていないWebアプリへのPOSTを拒否しました。');
      return ContentService.createTextOutput('Unauthorized');
    }

    if (e.postData && e.postData.contents) {
      const event = JSON.parse(e.postData.contents);
      if (event.type === 'MESSAGE' && event.message && event.message.text) {
        const messageText = event.message.text.trim();
        const senderName = event.message.sender && event.message.sender.displayName
          ? event.message.sender.displayName
          : '不明';
        const lines = messageText.split('\n');
        const title = lines[0];
        const content = lines.length > 1 ? lines.slice(1).join('\n') : '';

        const ss = SpreadsheetApp.openById(Config.SPREADSHEET_ID);
        const todoSheet = ss.getSheetByName(Config.SHEET_NAME_TODO);
        if (todoSheet) {
          todoSheet.appendRow([Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd"), 'チャット', title, content, '', '', '', '未着手', '', '', '', `送信者: ${senderName}`]);
          return ContentService.createTextOutput(JSON.stringify({ "text": `📝 ToDoに登録しました！\n件名: ${title}` })).setMimeType(ContentService.MimeType.JSON);
        }
      }
    }
  } catch (err) {
    console.error('doPost error:', err);
    return ContentService.createTextOutput('Bad Request');
  }
  return ContentService.createTextOutput("OK");
}

function admin_setupChatWebhookToken() {
  const props = PropertiesService.getScriptProperties();
  let token = props.getProperty(Config.CHAT_WEBHOOK_TOKEN_PROPERTY);
  if (!token) {
    token = Utilities.getUuid().replace(/-/g, '');
    props.setProperty(Config.CHAT_WEBHOOK_TOKEN_PROPERTY, token);
  }

  const serviceUrl = ScriptApp.getService().getUrl();
  const securedUrl = serviceUrl ? `${serviceUrl}?token=${encodeURIComponent(token)}` : '';
  const message = securedUrl
    ? `Google Chatの接続先URLを次に変更してください。\n\n${securedUrl}`
    : 'Webアプリを一度デプロイしてから、もう一度実行してください。';
  SpreadsheetApp.getUi().alert('🔐 チャット受信URLの設定', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

/* --- ヘルパー関数群 --- */
function getDashboardData(ss) {
  const dashSheet = ss.getSheetByName(Config.SHEET_NAME_DASHBOARD);
  if (!dashSheet) return { schedule: "予定なし", memo: "特になし" };
  let scheduleList = [];
  dashSheet.getRange(4, 2, 17, 4).getValues().forEach(row => {
    if (row[2] && row[2].toString().trim() !== "") {
      const s = row[0], e = row[1], memo = row[3];
      const sStr = (s instanceof Date) ? Utilities.formatDate(s, "JST", "HH:mm") : (s || "").toString();
      const eStr = (e instanceof Date) ? Utilities.formatDate(e, "JST", "HH:mm") : (e || "").toString();
      let timeLabel = (sStr === "終日" || sStr === "" || (sStr === "0:00" && (eStr === "" || eStr === "0:00"))) ? "【終日】" : `【${sStr}〜${eStr}】`;
      scheduleList.push(`・${timeLabel} ${row[2]} ${memo ? '（内容：' + memo + '）' : ''}`);
    }
  });
  const memosData = dashSheet.getRange(22, 2, 20, 1).getValues().filter(r => r[0] !== "");
  return { schedule: scheduleList.length > 0 ? scheduleList.join('\n') : "予定なし", memo: memosData.length > 0 ? memosData.map(m => `・${m[0]}`).join('\n') : "特になし" };
}

function getCompletedTasksToday(ss) {
  const data = ss.getSheetByName(Config.SHEET_NAME_TODO).getDataRange().getValues();
  data.shift(); 
  const todayStr = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd");
  let completedTasks = [];
  data.forEach(row => {
    if (row[Config.TODO_COL.STATUS] === '完了' && row[Config.TODO_COL.COMPLETED_DATE]) {
      try { if (Utilities.formatDate(new Date(row[Config.TODO_COL.COMPLETED_DATE]), "JST", "yyyy/MM/dd") === todayStr) completedTasks.push(`・${row[Config.TODO_COL.TITLE]}`); } catch(e) {}
    }
  });
  return completedTasks.length > 0 ? completedTasks.join('\n') : "本日完了した業務はありません。";
}

function askGeminiForReport(schedule, tasks, memo, modelName, dateHeader) {
  const prompt = `あなたは学校の教頭を支える秘書AIです。以下のデータから業務報告を作成してください。
【厳守事項】冒頭や末尾の挨拶は含めない。1行目は「**${dateHeader} 業務報告**」。構成は「**【主な動き（予定）】**」「**【業務実績（タスク）】**」「**【所感】**」の見出しのみ。所感は一般メモを教頭視点でビジネス文章にリライトすること。
【本日の予定】${schedule} \n【本日完了した業務】${tasks} \n【本日の気づき・一般メモ】${memo}`;
  const payload = { "contents": [{ "parts": [{ "text": prompt }] }] };
  try {
    const response = callGeminiWithRotation(payload, modelName);
    return JSON.parse(response.getContentText()).candidates[0].content.parts[0].text;
  } catch (e) {
    console.error("askGeminiForReport error:", e);
    return null;
  }
}

function sendChatNotification(masterSheet, docTitle, docUrl) {
  const webhookUrl = masterSheet.getRange(Config.MASTER_POS.WEBHOOK_CELL).getValue();
  if (webhookUrl && webhookUrl.startsWith('http')) {
    UrlFetchApp.fetch(webhookUrl, { "method": "post", "contentType": "application/json", "payload": JSON.stringify({ "text": `📝 **本日の教頭日誌（ドラフト）を作成しました！**\n\n📄 <${docUrl}|${docTitle}>` }) });
  }
}
