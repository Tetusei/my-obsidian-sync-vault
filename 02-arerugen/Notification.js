/**
 * 【ファイル9】Notification.gs (v4.1.3 日付文字列・完全対応版)
 * 役割：当日の給食対応情報をチャットに通知する
 * トリガー設定：毎朝 7:00〜8:00
 */
function sendDailyNotification() {
  // --- 👇ここから追加：土日スキップ処理👇 ---
  const today = new Date();
  const dayOfWeek = today.getDay(); // 0:日曜日 ～ 6:土曜日
  
  if (dayOfWeek === 0 || dayOfWeek === 6) {
    if (typeof writeLog === 'function') writeLog("💤 本日は土休日のため、チャット通知をスキップしました。", "info");
    return; // ここで処理を終了し、通知を送りません
  }
  // --- 👆追加ここまで👆 ---

  if (!CONFIG.isBotActive()) { console.log("💤 Bot機能OFFのため通知スキップ"); return; }

  const webhookUrl = CONFIG.getWebhookUrl();
  if (!webhookUrl) { writeLog("⚠️ Webhook URL未設定", "warn"); return; }

  const todayStr = Utilities.formatDate(today, "JST", "yyyy/MM/dd"); // 曜日判定で使ったtodayを再利用
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.VERIFY);
  
  if (!sheet || sheet.getLastRow() < CONFIG.START_ROW) return;

  const data = sheet.getRange(CONFIG.START_ROW, 1, sheet.getLastRow() - CONFIG.START_ROW + 1, 8).getValues();
  
  const todaysTargets = data.filter(row => {
    let dateVal = row[0];
    
    // 🌟 【修正】カレンダー同期と同じく、文字から日付を復元して「今日」か判定する
    if (typeof dateVal === 'string') {
      const match = dateVal.match(/^(\d{4}\/\d{2}\/\d{2})/);
      if (match) dateVal = new Date(match[1]);
    }
    
    return (dateVal instanceof Date) && !isNaN(dateVal.getTime()) && (Utilities.formatDate(dateVal, "JST", "yyyy/MM/dd") === todayStr);
  });

  let message = "";
  if (todaysTargets.length === 0) {
    message = `📅 ${todayStr}\n本日の給食アレルギー対応データはありません。`;
  } else {
    const menuName = todaysTargets[0][1];
    const details = todaysTargets.map(row => `・*${row[3]} ${row[4]}* (${row[5]}) : ${row[6]}`).join("\n");
    message = `🔔 *本日の給食アレルギー対応* (${todayStr})\n\n` +
              `🍛 *献立*: ${menuName}\n` +
              `⚠️ *対象者: ${todaysTargets.length}名*\n` +
              `${details}\n\n誤食のないよう確認をお願いします。`;
  }

  try {
    UrlFetchApp.fetch(webhookUrl, {
      "method": "post",
      "contentType": "application/json",
      "payload": JSON.stringify({ "text": message })
    });
    writeLog(`🔔 チャット通知完了: ${todayStr}`, "success");
  } catch (e) {
    writeLog(`❌ 通知エラー: ${e.message}`, "error");
  }
}