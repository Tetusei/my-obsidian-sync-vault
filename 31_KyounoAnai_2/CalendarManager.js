/**
 * CalendarManager.gs
 * VERSION: v48.6.0 (2026年度 運用安定版・WBGT自動取得機能追加)
 * ・Botシートへのカレンダー取込
 * ・分割登録機能（重複防止・アイコン付与・件数返却）
 * ・連絡・日課シートでチェックした行事をカレンダーから削除する機能
 * ・【NEW】環境省APIによるWBGT（暑さ指数）の自動取得
 */

/**
 * Botシートへのカレンダー取込（既存機能）
 */
function importCalendarToBotSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME_BOT);
  const settings = getSettings();
  const targetDate = sheet.getRange("A2").getValue();
  if (!(targetDate instanceof Date)) return;

  // 曜日・天気をセット
  sheet.getRange("B2").setValue(['日','月','火','水','木','金','土'][targetDate.getDay()] + "曜日");
  setDetailedWeather(sheet, targetDate);

  // エリアクリア
  const lastRow = Math.max(sheet.getLastRow(), CONFIG.SCHEDULE_START_ROW);
  sheet.getRange(CONFIG.SCHEDULE_START_ROW, 1, lastRow - CONFIG.SCHEDULE_START_ROW + 1, 7).clearContent().removeCheckboxes();

  let events = [];
  const rawIds = [settings.eventCalId, settings.menuCalId];
  const uniqueIds = rawIds.filter((id, index, self) => id && id.indexOf('@') > -1 && self.indexOf(id) === index);

  uniqueIds.forEach(id => {
    try {
      const cal = CalendarApp.getCalendarById(id);
      if (!cal) return;
      
      // 給食用カレンダーかどうかの判定
      const isMenuCal = (id === settings.menuCalId && id !== settings.eventCalId);

      cal.getEventsForDay(targetDate).forEach(ev => {
        const title = ev.getTitle();
        const isAllDay = ev.isAllDayEvent();

        // --- 優先スコアの決定 ---
        let priorityScore = 3; 
        if (isAllDay) {
          if (isMenuCal) {
            priorityScore = 1; // 献立（終日）を最優先
          } else {
            priorityScore = 2; // 行事（終日）を次点
          }
        }

        events.push({
          time: isAllDay ? "終日" : Utilities.formatDate(ev.getStartTime(), "JST", "HH:mm"),
          title: title,
          rawStart: isAllDay ? new Date(targetDate.getTime() - 1) : ev.getStartTime(),
          score: priorityScore
        });
      });
    } catch(e) { console.warn("カレンダー取得エラー: " + id); }
  });

  // --- スコアに基づいてソート ---
  events.sort((a, b) => {
    if (a.score !== b.score) return a.score - b.score;
    return a.rawStart - b.rawStart;
  });

  // シートへ書き込み
  if (events.length > 0) {
    const rows = events.map(e => [false, e.time, e.title, "", "", "", ""]);
    sheet.getRange(CONFIG.SCHEDULE_START_ROW, 1, rows.length, 7).setValues(rows);
    sheet.getRange(CONFIG.SCHEDULE_START_ROW, 1, rows.length, 1).insertCheckboxes();
  }
}

/**
 * 分割登録のバックエンド処理 (Admin.gsから呼び出される)
 */
function batchRegSplit(sheetName, mode, startRow, step) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  const settings = getSettings();
  
  if (!sheet) throw new Error(`シート「${sheetName}」が見つかりません。`);

  // カレンダーIDの特定
  const calId = (mode === 'Event') ? settings.eventCalId : settings.menuCalId;
  if (!calId) throw new Error(`カレンダーID未設定 (Mode: ${mode})`);
  
  const cal = CalendarApp.getCalendarById(calId);
  if (!cal) throw new Error(`カレンダー不可: ${calId}`);

  const lastRow = sheet.getLastRow();
  const endRow = Math.min(startRow + step - 1, lastRow);
  
  if (startRow > lastRow) {
    return { isFinished: true, nextRow: startRow, count: 0 };
  }

  const dataRange = sheet.getRange(startRow, 1, endRow - startRow + 1, 7);
  const values = dataRange.getValues();
  const results = [];
  const checks = [];
  let successCount = 0;

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const isChecked = row[0];
    const startDate = row[1];
    let title = row[2];
    const endDate = row[3];
    const isAllDay = row[4];
    const desc = row[5];
    const status = String(row[6]).trim();

    if (isChecked && status !== "済" && status !== "完了" && title && (startDate instanceof Date)) {
      try {
        if (mode === 'Event') {
          const options = { description: desc || "" };
          if (isAllDay) {
            cal.createAllDayEvent(title, startDate, options);
          } else {
            let validEndDate = (endDate instanceof Date) ? endDate : new Date(startDate.getTime() + 60 * 60 * 1000);
            if (validEndDate.getTime() <= startDate.getTime()) {
               validEndDate = new Date(startDate.getTime() + 60 * 60 * 1000);
            }
            cal.createEvent(title, startDate, validEndDate, options);
          }
        } else {
          title = getMenuIcon(title) + " " + title;
          cal.createAllDayEvent(title, startDate);
        }
        
        results.push(["済"]); 
        checks.push([false]); 
        successCount++; 

      } catch (e) {
        console.error(`Row ${startRow + i} Error: ${e.message}`);
        results.push(["エラー: " + e.message]);
        checks.push([true]); 
      }
    } else {
      results.push([row[6]]); 
      checks.push([row[0]]);
    }
  }

  sheet.getRange(startRow, 7, results.length, 1).setValues(results);
  sheet.getRange(startRow, 1, checks.length, 1).setValues(checks);

  return {
    isFinished: (endRow >= lastRow),
    nextRow: endRow + 1,
    count: successCount 
  };
}

function getMenuIcon(title) {
  if (!title) return "🍱";
  const t = title.toString();
  if (t.includes("カレー")) return "🍛";
  if (t.includes("パン")) return "🍞";
  if (t.includes("ごはん") || t.includes("ご飯") || t.includes("丼") || t.includes("飯")) return "🍚";
  if (t.includes("麺") || t.includes("スパゲティ") || t.includes("うどん") || t.includes("そば") || t.includes("パスタ")) return "🍜";
  if (t.includes("ケーキ") || t.includes("ゼリー") || t.includes("プリン")) return "🍰";
  if (t.includes("シチュー") || t.includes("スープ") || t.includes("汁")) return "🍲";
  if (t.includes("牛乳") || t.includes("ミルク")) return "🥛";
  return "🍱";
}

function deleteSelectedEvents() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME_BOT);
  const settings = getSettings();
  
  const targetDate = sheet.getRange("A2").getValue();
  if (!(targetDate instanceof Date)) {
    Browser.msgBox("エラー", "A2セルの日付が正しくありません。", Browser.Buttons.OK);
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.SCHEDULE_START_ROW) {
    Browser.msgBox("データなし", "削除対象のデータがありません。", Browser.Buttons.OK);
    return;
  }

  const dataRange = sheet.getRange(CONFIG.SCHEDULE_START_ROW, 1, lastRow - CONFIG.SCHEDULE_START_ROW + 1, 3);
  const values = dataRange.getValues();
  const targets = []; 

  values.forEach((row) => {
    if (row[0] === true && row[2]) { 
      targets.push(row[2].toString().trim());
    }
  });

  if (targets.length === 0) {
    Browser.msgBox("選択なし", "削除したい行事のチェックボックスをオンにしてください。", Browser.Buttons.OK);
    return;
  }

  const res = Browser.msgBox("削除確認", `チェックされた ${targets.length} 件の行事をカレンダーから削除しますか？\\n(カレンダーから完全に消えます)`, Browser.Buttons.OK_CANCEL);
  if (res === 'cancel') return;

  let deleteCount = 0;
  const calIds = [settings.eventCalId, settings.menuCalId].filter(id => id); 

  calIds.forEach(calId => {
    const cal = CalendarApp.getCalendarById(calId);
    if (!cal) return;

    const events = cal.getEventsForDay(targetDate);
    events.forEach(ev => {
      const evTitle = ev.getTitle().trim();
      const index = targets.indexOf(evTitle);
      if (index !== -1) {
        ev.deleteEvent();
        deleteCount++;
      }
    });
  });

  if (deleteCount > 0) {
    importCalendarToBotSheet();
    Browser.msgBox("完了", `${deleteCount} 件のイベントを削除し、最新の状態に更新しました。`, Browser.Buttons.OK);
  } else {
    Browser.msgBox("削除失敗", "該当するタイトルのイベントがカレンダーに見つかりませんでした。\\n(既に削除されているか、タイトルが一致しません)", Browser.Buttons.OK);
  }
}

function registerEventsFromBotSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME_BOT);
  const settings = getSettings();
  
  const targetDate = sheet.getRange("A2").getValue();
  if (!(targetDate instanceof Date)) {
    Browser.msgBox("エラー", "A2セルの日付が正しくありません。", Browser.Buttons.OK);
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.SCHEDULE_START_ROW) {
    Browser.msgBox("データなし", "登録対象の行がありません。", Browser.Buttons.OK);
    return;
  }

  const dataRange = sheet.getRange(CONFIG.SCHEDULE_START_ROW, 1, lastRow - CONFIG.SCHEDULE_START_ROW + 1, 4);
  const values = dataRange.getValues();
  const checks = []; 

  let count = 0;
  const cal = CalendarApp.getCalendarById(settings.eventCalId); 

  if (!cal) {
    Browser.msgBox("エラー", "行事用カレンダーが見つかりません。", Browser.Buttons.OK);
    return;
  }

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const isChecked = row[0];  
    const timeStr = String(row[1]); 
    const title = String(row[2]);   
    const desc = String(row[3]);    

    if (isChecked && title) {
      try {
        if (timeStr.includes("終日") || timeStr === "") {
          cal.createAllDayEvent(title, targetDate, { description: desc });
        } else {
          const timeParts = timeStr.match(/(\d{1,2})[:：](\d{2})/);
          if (timeParts) {
            const startDate = new Date(targetDate);
            startDate.setHours(parseInt(timeParts[1]), parseInt(timeParts[2]), 0);
            const endDate = new Date(startDate.getTime() + 60 * 60 * 1000); 
            cal.createEvent(title, startDate, endDate, { description: desc });
          } else {
             cal.createAllDayEvent(title, targetDate, { description: desc + "\n(時間不明のため終日登録)" });
          }
        }
        
        count++;
        checks.push([false]); 
      } catch (e) {
        console.error(e);
        checks.push([true]); 
      }
    } else {
      checks.push([row[0]]); 
    }
  }

  sheet.getRange(CONFIG.SCHEDULE_START_ROW, 1, checks.length, 1).setValues(checks);

  if (count > 0) {
    Browser.msgBox("登録完了", `${count} 件の行事をカレンダーに追加しました。`, Browser.Buttons.OK);
  } else {
    Browser.msgBox("対象なし", "チェックされた行事が見つかりませんでした。", Browser.Buttons.OK);
  }
}

function updateAreaCode() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ss.getSheetByName("基礎データ"); 
  
  if (!targetSheet) {
    Browser.msgBox("エラー", "「基礎データ」シートが見つかりません。シート名を確認してください。", Browser.Buttons.OK);
    return;
  }

  const placeName = targetSheet.getRange("B12").getValue().toString().trim();
  if (!placeName) {
    Browser.msgBox("エラー", "「基礎データ」シートのB12セルに地域名（例：波佐見町）を入力してください。", Browser.Buttons.OK);
    return;
  }

  const areaMap = {
    "波佐見": "420022", "波佐見町": "420022",
    "佐世保": "420020", "佐世保市": "420020",
    "長崎":   "420010", "長崎市":   "420010",
    "佐賀":   "410010", "有田":     "410020"
  };

  let code = areaMap[placeName];
  if (!code) {
    for (let key in areaMap) {
      if (placeName.indexOf(key) > -1) {
        code = areaMap[key];
        break;
      }
    }
  }

  if (code) {
    targetSheet.getRange("C12").setValue(code);
    Browser.msgBox("完了", `「基礎データ」シートの地域「${placeName}」を判定し、\nコード（${code}）をC12セルにセットしました。`, Browser.Buttons.OK);
  } else {
    Browser.msgBox("未登録", `「${placeName}」に対応するコードが登録されていません。\nスクリプト内の areaMap に追加してください。`, Browser.Buttons.OK);
  }
}

/**
 * 気象庁APIから詳細な天気を取得し、環境省からWBGTを取得してD2に書き込む
 */
function setDetailedWeather(sheet, date) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("基礎データ");
  
  const targetStr = (date instanceof Date) ? Utilities.formatDate(date, "JST", "yyyyMMdd") : "";

  if (!targetStr) {
    sheet.getRange("D2").setValue("取得失敗");
    return;
  }

  try {
    sheet.getRange("D2").setValue(""); 
    const targetDateStr = Utilities.formatDate(date, "JST", "yyyy-MM-dd");
    let areaCode = configSheet.getRange('C12').getValue(); 
    if (!areaCode) return;

    // --- 1. 気象庁APIから天気の取得 ---
    let json = null;
    let currentCode = areaCode;
    const MAX_RETRY = 3; 

    for (let i = 0; i < MAX_RETRY; i++) {
      try {
        const url = `https://www.jma.go.jp/bosai/forecast/data/forecast/${currentCode}.json`;
        const res = UrlFetchApp.fetch(url);
        if (res.getResponseCode() === 200) { 
          json = JSON.parse(res.getContentText()); 
          break; 
        }
      } catch (e) {
        try {
          const areaRes = UrlFetchApp.fetch("https://www.jma.go.jp/bosai/common/const/area.json");
          const areaData = JSON.parse(areaRes.getContentText());
          let nextCode = null;
          if (areaData.class20s && areaData.class20s[currentCode]) nextCode = areaData.class20s[currentCode].parent;
          else if (areaData.class15s && areaData.class15s[currentCode]) nextCode = areaData.class15s[currentCode].parent;
          else if (areaData.class10s && areaData.class10s[currentCode]) nextCode = areaData.class10s[currentCode].parent;
          if (nextCode && nextCode !== currentCode) currentCode = nextCode; 
          else currentCode = currentCode.toString().substring(0, 2) + "0000"; 
        } catch (ex) { break; }
      }
    }

    if (!json) { sheet.getRange("D2").setValue("取得失敗"); return; }
    
    const daily = json[0];
    let weatherStr = "";
    let targetAreaIdx = 0;
    const areas = daily.timeSeries[0].areas;
    for (let k = 0; k < areas.length; k++) {
      if (areas[k].area.code == areaCode || areas[k].area.code == currentCode) { targetAreaIdx = k; break; }
    }

    const wDates = daily.timeSeries[0].timeDefines;
    for (let i = 0; i < wDates.length; i++) {
      if (Utilities.formatDate(new Date(wDates[i]), "JST", "yyyy-MM-dd") === targetDateStr) {
        let w = daily.timeSeries[0].areas[targetAreaIdx].weathers[i];
        w = w.replace(/ /g, " ").replace(/晴れ?/g, "☀").replace(/(くもり|曇り?)/g, "☁")
             .replace(/雨/g, "☔").replace(/雪/g, "☃").replace(/雷/g, "⚡")
             .replace(/のち/g, "→").replace(/時々/g, "｜").replace(/所により/g, "ᴾ");
        weatherStr += w; break;
      }
    }

    let tempStr = "";
    try {
      const tDates = daily.timeSeries[2].timeDefines;
      let dayTemps = [];
      for(let i=0; i<tDates.length; i++) {
        if (tDates[i].startsWith(targetDateStr)) {
          let tAreas = daily.timeSeries[2].areas;
          let tAreaIdx = 0;
          for(let m=0; m<tAreas.length; m++) { if(tAreas[m].area.code == areaCode || tAreas[m].area.code == currentCode) tAreaIdx = m; }
          let t = tAreas[tAreaIdx].temps[i]; 
          if(t!==null && t!=="") dayTemps.push(t);
        }
      }
      if (dayTemps.length >= 2) tempStr = " " + Math.max(...dayTemps) + "℃/" + Math.min(...dayTemps) + "℃";
      else if (dayTemps.length === 1) tempStr = " " + dayTemps[0] + "℃";
    } catch(e) {}

    let rainStr = "";
    try {
      const pDates = daily.timeSeries[1].timeDefines;
      let rainProbs = [];
      for (let i = 0; i < pDates.length; i++) {
        if (pDates[i].startsWith(targetDateStr)) { 
          rainProbs.push(parseInt(daily.timeSeries[1].areas[targetAreaIdx].pops[i])); 
        }
      }
      if (rainProbs.length > 0) rainStr = " " + Math.max(...rainProbs) + "%";
    } catch(e) {}

    // --- 2. 環境省データからWBGTを取得 (佐世保: 84366) ---
    let wbgtStr = getWbgtString(targetDateStr);

    // --- 3. D2セルに書き込み ---
    sheet.getRange("D2").setValue(weatherStr + tempStr + rainStr + wbgtStr);

  } catch (e) { 
    sheet.getRange("D2").setValue("取得失敗"); 
  }
}

/**
 * 環境省CSVから「今日」のWBGT最大値と、その「時刻」をピンポイントで取得する
 */
function getWbgtString(dateStr) {
  const urls = [
    "https://www.wbgt.env.go.jp/prev15WG/dl/yohou_84266.csv", // 先生が見ている観測所
    "https://www.wbgt.env.go.jp/prev15WG/dl/yohou_nagasaki.csv",
    "https://www.wbgt.env.go.jp/prev15WG/dl/yohou_all.csv"
  ];
  
  let csvData = null;
  for (let url of urls) {
    try {
      const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      if (res.getResponseCode() === 200) {
        const text = res.getContentText("Shift_JIS");
        if (!text.includes("<html") && !text.includes("<HTML")) {
          csvData = text;
          break; 
        }
      }
    } catch(e) {}
  }

  if (!csvData) return " (WBGT: 通信エラー)"; 

  try {
    const lines = csvData.split("\n");
    let targetCols = [];
    let colToHour = {}; // ★列番号と「時刻」を紐づけて記憶する辞書
    let maxWbgt = -1;
    let maxTime = "";   // ★最大値が出た時刻を記録する変数

    // 1. ヘッダー行を探し、「今日」の列番号と「時刻」をすべて記憶する
    for (let i = 0; i < lines.length; i++) {
      let cols = lines[i].split(",");
      let isHeader = false;
      let firstDateStr = null;
      
      for (let c = 0; c < cols.length; c++) {
         let val = cols[c].replace(/"/g, "").trim();
         // 「2026060509」のような10桁の数字（時刻データ）があればヘッダー
         if (val.length === 10 && val.startsWith("202")) {
            isHeader = true;
            let colDate = val.substring(0, 8); // "20260605" の部分
            let colHour = parseInt(val.substring(8, 10), 10); // "09" を 9 に変換
            
            if (!firstDateStr) {
               firstDateStr = colDate; // 一番最初に出てきた日付を「今日」と確定する
            }
            // その「今日」と同じ日付を持つ列の番号と時刻をセットで記録！
            if (colDate === firstDateStr) {
               targetCols.push(c);
               colToHour[c] = colHour + "時"; // 例：14時
            }
         }
      }
      if (isHeader) break;
    }

    // 2. 84266 の行を探し、記憶した「今日の列」だけを見て最大値を探す
    for (let i = 0; i < lines.length; i++) {
      if (lines[i].includes("84266")) { 
        const cols = lines[i].split(",");
        
        if (targetCols.length > 0) {
          for (let c of targetCols) {
            if (c < cols.length) {
               let valStr = (cols[c] || "").replace(/"/g, "").trim();
               if (valStr === "" || isNaN(valStr)) continue;
               
               let val = parseFloat(valStr);
               if (val > 0) {
                 if (val > 100) val = val / 10; // 217 -> 21.7 
                 
                 // ★今の最大値よりも高い数値が出たら、数値と時刻を両方更新する！
                 if (val < 45 && val > maxWbgt) {
                   maxWbgt = val;
                   maxTime = colToHour[c]; 
                 }
               }
            }
          }
        }
        break; // 84266の行の処理が終わったら終了
      }
    }

    if (maxWbgt === -1) return " (WBGT: データなし)"; 
    
    // 小数点第1位まで表示
    const displayWbgt = maxWbgt.toFixed(1);
    const checkVal = Math.round(maxWbgt);
    
    let level = "ほぼ安全";
    if (checkVal >= 31) level = "危険";
    else if (checkVal >= 28) level = "厳重警戒";
    else if (checkVal >= 25) level = "警戒";
    else if (checkVal >= 21) level = "注意";
    
    // ★時刻をカッコに入れて一緒に表示するように修正しました
    return ` (🥵WBGT最大: ${displayWbgt} [${maxTime}] ${level})`;

  } catch(e) {
    return " (WBGT: 内部エラー)";
  }
}

/**
 * 11時〜12時の間に実行されるタイマートリガーを全自動でセットする関数
 */
function setupWbgtTrigger() {
  // 二重登録を防ぐため、既存のトリガーを一度きれいに削除します
  clearWbgtTrigger();
  
  // 11時〜12時の間に動くタイマーをセット
  ScriptApp.newTrigger('sendCurrentWbgtNotification')
    .timeBased()
    .atHour(11) // 11時台（11:00〜12:00の間）に実行
    .everyDays(1)
    .create();
    
  SpreadsheetApp.getUi().alert('設定完了：毎日11時〜12時の間に自動通知するトリガーをセットしました！');
}

/**
 * セットされているWBGT用の自動トリガーをすべて解除（削除）する関数
 */
function clearWbgtTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  let count = 0;
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'sendCurrentWbgtNotification') {
      ScriptApp.deleteTrigger(triggers[i]);
      count++;
    }
  }
  if (count > 0) {
    Logger.log(count + "個のトリガーを削除しました。");
  }
}

/**
 * 基礎データのON/OFFを確認し、Open-Meteo APIから波佐見町の最新の気温・湿度を取得して
 * 現在のリアルタイムWBGTを計算し、指定のWebhookへ通知する
 */
function sendCurrentWbgtNotification() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingSheet = ss.getSheetByName("基礎データ");
  if (!settingSheet) return;

  // 1. B13セルのON/OFFチェック
  const isEnabled = String(settingSheet.getRange("B13").getValue()).trim().toUpperCase();
  if (isEnabled !== "ON") {
    Logger.log("WBGT個別通知はOFFのためスキップします。");
    return; 
  }

  // 2. B14セルのWebhook URL取得
  const webhookUrl = String(settingSheet.getRange("B14").getValue()).trim();
  if (!webhookUrl || !webhookUrl.startsWith("http")) {
    Logger.log("有効なWebhook URLがB14セルに設定されていません。");
    return;
  }

  // 波佐見町の座標
  const HASAMI_LAT = 33.1456;
  const HASAMI_LON = 129.9063;

  try {
    // 3. Open-Meteo APIからリアルタイムデータを取得
    const apiUrl = `https://api.open-meteo.com/v1/forecast?latitude=${HASAMI_LAT}&longitude=${HASAMI_LON}&current=temperature_2m,relativehumidity_2m&timezone=Asia%2FTokyo`;
    
    const response = UrlFetchApp.fetch(apiUrl);
    const json = JSON.parse(response.getContentText());

    // 4. 現在の気温と湿度を取得
    const currentData = json.current;
    const temp = currentData.temperature_2m; 
    const humidity = currentData.relativehumidity_2m; 

    // 5. 計算式を使ってWBGT（推定値）を計算
    const wbgtEst = (0.735 * temp) + (0.0374 * humidity) + (0.00292 * temp * humidity) - 4.064;
    const wbgtFormatted = Math.round(wbgtEst * 10) / 10;

    // 警戒レベルを判定
    const checkVal = Math.round(wbgtFormatted);
    let level = "ほぼ安全";
    if (checkVal >= 31) level = "危険";
    else if (checkVal >= 28) level = "厳重警戒";
    else if (checkVal >= 25) level = "警戒";
    else if (checkVal >= 21) level = "注意";

    // 通知用の時間文字列を作成（日本時間）
    const now = new Date();
    const timeString = Utilities.formatDate(now, 'Asia/Tokyo', 'H時m分');

    // 6. チャットへの通知メッセージを作成
    const message = `【現在のリアルタイムWBGT】\n状態: 🟢波佐見町実測値より算出\n時刻: ${timeString}現在\n気温: ${temp}℃ / 湿度: ${humidity}%\n計算WBGT: 🥵${wbgtFormatted.toFixed(1)} [${level}]`;

    // 7. Webhookで指定 of チャットへ送信
    const postOptions = {
      "method": "post",
      "contentType": "application/json",
      "payload": JSON.stringify({ "text": message })
    };
    
    UrlFetchApp.fetch(webhookUrl, postOptions);
    Logger.log(`現在値を正常に送信しました。${timeString}現在 WBGT:${wbgtFormatted}`);

  } catch (e) {
    Logger.log(`通知処理内でエラー: ${e.message}`);
  }
}