/**
 * Main.gs
 * VERSION: v49.3 (重複排除・完全整合版)
 * ・トリガー実行、手動実行、日付変更アクションの集約
 */

// ==========================================
//  1. Bot送信 (自動・手動)
// ==========================================

/**
 * 毎日決まった時間にBot送信する（時間主導型トリガー用）
 */
function sendDailyChatNotification() {
  const today = new Date();
  const day = today.getDay(); // 0: 日曜日, 6: 土曜日
  if (day === 0 || day === 6) {
    Logger.log("土曜日・日曜日のため、送信をスキップしました。");
    return;
  }

  const settings = getSettings();
  if (!settings.isBotEnabled) return;
  
  // まず日付を今日にして更新
  forceUpdateToToday();
  
  // 送信
  constructAndSendChat(false); 
}

/**
 * 手動送信ボタン・メニュー用
 * (Admin.gsのメニューからも、この関数が呼ばれます)
 */
function sendManualChat() {
  constructAndSendChat(true); 
}

// ==========================================
//  2. 自動更新トリガー
// ==========================================

/**
 * シート編集時に動作するトリガー
 */
function respondToEdit(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = e.source.getActiveSheet();
  
  // 連絡・日課シートのA2セル（日付）が変更されたら再読込
  if (sheet.getName() === CONFIG.SHEET_NAME_BOT && e.range.getA1Notation() === "A2") {
    ss.toast("🔄 カレンダーと天気を取得しています...", "読み込み中", -1);
    importCalendarToBotSheet(); // [CalendarManager.gs]
    ss.toast("✅ 日付の更新が完了しました！", "完了", 3);
  }
}

// ==========================================
//  3. 日付操作アクション (ボタン・メニュー共通)
// ==========================================

/**
 * 日付を「今日」にしてカレンダーを取り込む
 * (Admin.gs, Ai_Pdf.gs からも呼ばれます)
 */
function forceUpdateToToday() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME_BOT);
  if (!sheet) return;

  sheet.getRange("A2").setValue(new Date());
  
  // カレンダー再読み込み
  importCalendarToBotSheet(); 
  // ついでに天気も更新
  if (typeof setDetailedWeather === 'function') {
    setDetailedWeather(sheet, new Date());
  }
}

/**
 * 日付を「今日」にする (forceUpdateToTodayのエイリアス)
 * ※ボタン割り当てなどで名前が変わっても良いように残します
 */
function moveToToday() {
  forceUpdateToToday();
}