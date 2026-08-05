/**
 * BotNotification.gs
 * VERSION: v49.6 (天気取得機能をCalendarManagerへ完全移行)
 * メッセージ構築、送信ロジックのみを担当
 */

/**
 * メッセージを構築し、Google Chatへ送信する
 */
function constructAndSendChat(isManual) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME_BOT);
  const kisoSheet = ss.getSheetByName("基礎データ"); 
  const settings = getSettings();
  
  // A2セルの日付取得
  const dateVal = sheet.getRange("A2").getValue();
  const dateStr = (dateVal instanceof Date) ? Utilities.formatDate(dateVal, "JST", "MM/dd(E)") : "日付不明";
  
  const todayStr = Utilities.formatDate(new Date(), "JST", "yyyyMMdd");
  const targetDateStr = (dateVal instanceof Date) ? Utilities.formatDate(dateVal, "JST", "yyyyMMdd") : "";
  const isToday = (todayStr === targetDateStr);

  // --- 1. 行事カウントダウンの作成 ---
  let countMsg = "";
  const eventName = kisoSheet.getRange("B9").getValue();    // 行事名
  const eventDate = kisoSheet.getRange("B10").getValue();   // 行事の日付
  
  if (eventName && eventDate instanceof Date && dateVal instanceof Date) {
    const diff = Math.ceil((eventDate.setHours(0,0,0,0) - dateVal.setHours(0,0,0,0)) / 86400000);
    if (diff > 0) {
      countMsg = `\n⏳ *${eventName}まで あと${diff}日*`;
    } else if (diff === 0) {
      countMsg = `\n🎉 *本日は ${eventName} です！*`;
    }
  }

  // --- 2. 天気情報の作成 ---
  let weatherSection = "";
  const weather = sheet.getRange("D2").getValue() || "";
  if (weather && String(weather) !== "取得失敗") {
    weatherSection = `🌍 *天気*: ${weather}\n`; 
  }

  // --- 3. お知らせ等の取得 ---
  const notice = sheet.getRange("C2").getValue();

  // --- 4. スケジュールと給食の分離・作成 ---
  let lunchLines = [];    // 給食格納用
  let scheduleLines = []; // その他の予定格納用
  let scheduleText = "・特になし";

  // 給食判定用のキーワード
  const keywords = (CONFIG.LUNCH_KEYWORDS && CONFIG.LUNCH_KEYWORDS.length > 0) 
    ? CONFIG.LUNCH_KEYWORDS 
    : ['ごはん', 'パン', '麺', 'カレー', 'シチュー', '給食', 'ランチ', '牛乳', '🍞', '🍚', '🍛', '🍱', '🥛', '🍝', '🥪'];

  const lastRow = sheet.getLastRow();
  if (lastRow >= CONFIG.SCHEDULE_START_ROW) {
    const data = sheet.getRange(CONFIG.SCHEDULE_START_ROW, 1, lastRow - CONFIG.SCHEDULE_START_ROW + 1, 4).getValues();
    
    data.forEach(r => {
      const timeVal = r[1];
      const content = r[2];
      
      // 空白または「健康観察」は除外
      if (!content || String(content).includes("健康観察")) return;

      const strContent = String(content);
      const isLunch = keywords.some(k => strContent.includes(k));

      if (isLunch) {
        lunchLines.push(`・${content}`);
      } else {
        let time = (timeVal instanceof Date) ? Utilities.formatDate(timeVal, "JST", "HH:mm") : timeVal;
        const line = (time === "終日" || time === "") ? `・${content}` : `・${time} ${content}`;
        scheduleLines.push(line);
      }
    });

    if (scheduleLines.length > 0) scheduleText = scheduleLines.join("\n");
  }

  // --- 5. AIコメントの作成 ---
  let aiComment = "";
  const aiStatus = kisoSheet.getRange("B11").getValue(); // ON/OFF
  if (isToday && aiStatus === "ON") {
    const prompt = `あなたは学校のAIです。日付:${dateStr}, 予定:${scheduleText.substring(0,100)}。毎朝送信されるメッセージですので、先生方への朝の挨拶と、今日一日を応援する前向きな一言を80文字以内で作成してください。過去形や「お疲れ様でした」は使用せず、朝にふさわしい表現にしてください。`;
    try {
      const aiText = callGeminiWithRetry(CONFIG.MODEL_NAME, prompt, settings);
      if (aiText) {
        aiComment = `\n\n🤖 *AIより*\n${aiText}`;
      }
    } catch(e) {
      console.warn("AIコメント生成失敗: " + e.message);
    }
  }

  // --- 6. メッセージの組み立て ---
  let lunchSection = "";
  if (lunchLines.length > 0) {
    lunchSection = `🍱 *本日の給食*\n${lunchLines.join("\n")}\n\n`;
  }

  const noticeSection = (isManual && notice) ? `📢 *お知らせ*\n${notice}\n\n` : "";
  
  const message = `📅 *${dateStr} の予定*${countMsg}\n${weatherSection}${lunchSection}${noticeSection}📝 *当日の日程*\n${scheduleText}${aiComment}`;

  // --- 7. 送信処理 ---
  try {
    UrlFetchApp.fetch(settings.webhookUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ text: message })
    });
    if (isManual) ss.toast("Google Chatへ送信しました");
  } catch(e) {
    if (isManual) Browser.msgBox("送信エラー", e.message, Browser.Buttons.OK);
  }
}