/**
 * 【ファイル8】Calendar.gs (v4.1.3 日付文字列・完全対応版)
 * 役割：VerifyシートのデータをGoogleカレンダーに同期する
 */
function syncVerifyToCalendar() {
  if (!CONFIG.isBotActive()) {
    writeLog("💤 Bot機能OFFのためカレンダー同期をスキップ", "info");
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const verifySheet = ss.getSheetByName(CONFIG.SHEET_NAMES.VERIFY);
  
  if (!verifySheet || verifySheet.getLastRow() < CONFIG.START_ROW) {
    writeLog("⚠️ カレンダー同期スキップ: データなし", "warn");
    return;
  }

  writeLog("🚀 カレンダー同期を開始します...", "info");

  // 1. データ集計
  const data = verifySheet.getRange(CONFIG.START_ROW, 1, verifySheet.getLastRow() - CONFIG.START_ROW + 1, 8).getValues();
  const scheduleMap = new Map();

  data.forEach(row => {
    let dateVal = row[0];
    if (!dateVal) return;
    
    // 🌟 【修正】文字（例: "2026/03/13(金)"）からカレンダー用の日付を復元する
    if (typeof dateVal === 'string') {
      const match = dateVal.match(/^(\d{4}\/\d{2}\/\d{2})/);
      if (match) {
        dateVal = new Date(match[1]);
      }
    }
    
    // それでも日付として認識できない行はスキップ
    if (!(dateVal instanceof Date) || isNaN(dateVal.getTime())) return;

    const dateKey = Utilities.formatDate(dateVal, "JST", "yyyy/MM/dd");
    
    // リスト表記: ・1-A 山田 (卵) : 除去食
    const personInfo = `・【${row[3]} ${row[4]}】 (${row[5]}) : ${row[6]}`;
    
    if (!scheduleMap.has(dateKey)) {
      scheduleMap.set(dateKey, { date: dateVal, menu: row[1], details: [] });
    }
    scheduleMap.get(dateKey).details.push(personInfo);
  });

  // 2. カレンダー登録
  const calendarId = CONFIG.getCalendarId();
  const calendar = CalendarApp.getCalendarById(calendarId);
  if (!calendar) { writeLog(`❌ カレンダー不明: ${calendarId}`, "error"); return; }

  let updateCount = 0;
  const EVENT_PREFIX = '🍱給食:'; 

  scheduleMap.forEach((val, key) => {
    // 重複削除（洗い替え）
    const events = calendar.getEventsForDay(val.date);
    events.forEach(e => { if (e.getTitle().startsWith(EVENT_PREFIX)) e.deleteEvent(); });

    // 新規登録
    const title = `${EVENT_PREFIX} ${val.details.length}名対応`;
    const description = `【献立】${val.menu}\n\n【対応リスト】\n${val.details.join("\n")}\n\n※自動生成`;
    calendar.createAllDayEvent(title, val.date, { description: description });
    updateCount++;
  });

  writeLog(`✅ カレンダー同期完了: ${updateCount}日分更新`, "success");
}