/**
 * Main.gs
 * 役割：ダッシュボードの表示更新、およびデータベース(Main)への追記
 */

function refreshDashboardFormulas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.DASHBOARD);
  
  // --- 1. まず B3（献立名）をリフレッシュ ---
  const b3 = dashSheet.getRange("B3");
  const f3 = b3.getFormula();
  if (f3) {
    b3.setFormula("");
    SpreadsheetApp.flush(); 
    b3.setFormula(f3);
  }

  // --- 2. 献立名がセルに書き込まれるのを確実に待つ ---
  Utilities.sleep(500); 
  SpreadsheetApp.flush();

  // --- 3. 次に B4（材料）を「強力に」叩き直す ---
  const b4 = dashSheet.getRange("B4");
  const f4 = b4.getFormula();
  if (f4) {
    b4.setFormula("");
    SpreadsheetApp.flush(); 
    b4.setFormula(f4);
  }

  // --- 4. 最後に下のリスト（B8:E30）をGASで更新 ---
  updateDashboardListByGas(dashSheet);
}

function updateDashboardListByGas(dashSheet) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mainSheetName = (typeof CONFIG !== 'undefined' && CONFIG.SHEET_NAMES && CONFIG.SHEET_NAMES.MAIN) ? CONFIG.SHEET_NAMES.MAIN : "Main";
    const masterSheetName = (typeof CONFIG !== 'undefined' && CONFIG.SHEET_NAMES && CONFIG.SHEET_NAMES.MASTER) ? CONFIG.SHEET_NAMES.MASTER : "献立マスタ";
    
    const mainSheet = ss.getSheetByName(mainSheetName); 
    const masterSheet = ss.getSheetByName(masterSheetName); 
    const dbSheetObj = ss.getSheetByName("AllDB");
    if (!dbSheetObj) {
      if (typeof writeLog === 'function') writeLog("⚠️ AllDBシートが見つかりません。", "error");
      return;
    }

    // 🌟ダッシュボード更新前に、AllDBを最新に同期（Mainの編集内容を取り込む）
    let syncResult = "Pending";
    if (mainSheet && masterSheet && typeof saveToAllDB === 'function') {
      try {
        saveToAllDB(mainSheet, masterSheet, dbSheetObj);
        syncResult = "Success";
      } catch (e) {
        syncResult = `Error: ${e.message}`;
        if (typeof writeLog === 'function') writeLog(`⚠️ AllDBの同期中にエラーが発生しました: ${e.message}`, "error");
      }
    } else {
      syncResult = `Skipped: main=${!!mainSheet}, master=${!!masterSheet}`;
      if (typeof writeLog === 'function') {
        writeLog(`⚠️ 同期スキップ: main=${!!mainSheet}, master=${!!masterSheet}, db=${!!dbSheetObj}`, "warn");
      }
    }

    // ダッシュボードの対象日を取得
    const targetDateVal = dashSheet.getRange(CONFIG.DASHBOARD.DATE_CELL).getValue();
    if (!targetDateVal) return;
    
    // 対象日を yyyy/MM/dd の文字列に安全に変換
    let targetDateStr = "";
    if (targetDateVal instanceof Date) {
      targetDateStr = Utilities.formatDate(targetDateVal, "JST", "yyyy/MM/dd");
    } else {
      const m = String(targetDateVal).match(/^(\d{4})[\/\-年](\d{1,2})[\/\-月](\d{1,2})/);
      if (m) {
        targetDateStr = m[1] + "/" + ("0" + m[2]).slice(-2) + "/" + ("0" + m[3]).slice(-2);
      } else {
        const d = new Date(targetDateVal);
        if (!isNaN(d.getTime())) {
          targetDateStr = Utilities.formatDate(d, "JST", "yyyy/MM/dd");
        }
      }
    }
    if (!targetDateStr) return;

    // AllDBシートからデータを取得
    let dbData = dbSheetObj.getDataRange().getValues();
    if (dbData.length <= 1) {
      clearDashboardList_(dashSheet);
      return;
    }
    
    dbData.shift(); // ヘッダー行を除外

    // 🌟 基礎データのB10セルから除外キーワードを取得
    let excludeKeywords = [];
    const baseSheet = ss.getSheetByName(mainSheetName === "Main" ? "基礎データ" : (CONFIG.SHEET_NAMES.CONFIG || "基礎データ"));
    if (baseSheet) {
      const excludeValue = baseSheet.getRange("B10").getValue();
      if (excludeValue) {
        excludeKeywords = String(excludeValue).split(/[、,，\s\n\r]+/).map(k => k.trim()).filter(Boolean);
      }
    }

    const todaysData = dbData.filter(row => {
      const dateCell = row[0]; // A列: 日付
      if (!dateCell) return false;
      
      let rowDateStr = "";
      if (dateCell instanceof Date) {
        rowDateStr = Utilities.formatDate(dateCell, "JST", "yyyy/MM/dd");
      } else {
        const m = String(dateCell).trim().match(/^(\d{4})[\/\-年](\d{1,2})[\/\-月](\d{1,2})/);
        if (m) {
          rowDateStr = m[1] + "/" + ("0" + m[2]).slice(-2) + "/" + ("0" + m[3]).slice(-2);
        }
      }
      if (rowDateStr !== targetDateStr) return false;

      // 🌟B10セルの除外キーワードに一致する料理名があれば除外
      const actionText = String(row[6] || ""); // G列: 対応内容
      const matchDish = actionText.match(/^【(.*?)】/);
      const dishName = matchDish ? matchDish[1].trim() : "";
      if (dishName && excludeKeywords.length > 0) {
        const isExcluded = excludeKeywords.some(keyword => dishName.includes(keyword));
        if (isExcluded) return false; 
      }

      return true;
    });


    if (typeof writeLog === 'function') {
      const debugTime = Utilities.formatDate(new Date(), "JST", "HH:mm:ss");
      const debugText = `[Debug] ${debugTime} | 対象日: ${targetDateStr} | AllDB総数: ${dbData.length}件 | 抽出数: ${todaysData.length}件 | sync=${syncResult}`;
      writeLog(debugText, "info");
      if (todaysData.length > 0) {
        const names = todaysData.map(r => r[4]).join(", ");
        writeLog(`[Debug] 抽出された生徒: ${names}`, "info");
      }
    }

    // --- 集約処理（同じ名前・クラスの生徒を1つに合体） ---
    const groupedMap = {};
    todaysData.forEach(row => {
      const className = String(row[3] || "").trim();    // D列: クラス
      const studentName = String(row[4] || "").trim();  // E列: 氏名
      const groupKey = className + "_" + studentName;

      const actionText = String(row[6] || "").trim();   // G列: 対応内容

      if (!groupedMap[groupKey]) {
        groupedMap[groupKey] = {
          className: className,
          studentName: studentName,
          allergen: String(row[5] || ""),               // F列: アレルゲン
          actions: actionText ? [actionText] : []
        };
      } else {
        const current = groupedMap[groupKey];
        if (actionText && !current.actions.includes(actionText)) {
          current.actions.push(actionText); 
        }
        if (row[5]) {
          const currentAllergens = current.allergen.split(/[、,，\s]+/).map(a => a.trim()).filter(Boolean);
          const newAllergens = String(row[5]).split(/[、,，\s]+/).map(a => a.trim()).filter(Boolean);
          current.allergen = [...new Set([...currentAllergens, ...newAllergens])].join("、");
        }
      }
    });

    // 🌟B〜E列の4列分だけの配列にする（F列・G列には一切触れない）
    const groupedList = Object.keys(groupedMap).map(key => {
      const g = groupedMap[key];
      return [g.className, g.studentName, g.allergen, g.actions.join("\n")];
    });

    groupedList.sort((a, b) => {
      if (a[0] !== b[0]) return a[0] < b[0] ? -1 : 1; 
      return a[1] < b[1] ? -1 : 1; 
    });

    clearDashboardList_(dashSheet);

    if (groupedList.length > 0) {
      // 🌟B〜E列（4列分）だけを一括書き込み
      dashSheet.getRange(8, 2, groupedList.length, 4).setValues(groupedList);
    }

    // E2セルの「対象者数」を更新
    dashSheet.getRange(CONFIG.DASHBOARD.TARGET_COUNT_CELL).setValue(`${groupedList.length}名`);
  } catch (err) {
    if (typeof writeLog === 'function') writeLog(`❌ updateDashboardListByGasエラー: ${err.message}`, "error");
    throw err;
  }
}

function clearDashboardList_(dashSheet) {
  // 🌟B〜E列の文字を最終行まで動的にクリア（30行目以降の幽霊データを残さないため）
  const lastRow = dashSheet.getLastRow();
  if (lastRow >= 8) {
    const range = dashSheet.getRange(8, 2, lastRow - 7, 4); // B8:E[lastRow]
    range.clearContent();
  }
}

function dailyUpdateForGoogleSites() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.DASHBOARD);
  
  if (!dashSheet) return;

  const todayFormatted = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd");
  
  dashSheet.getRange(CONFIG.DASHBOARD.DATE_CELL).setValue(todayFormatted);
  SpreadsheetApp.flush();
  
  refreshDashboardFormulas();
  
  console.log("日付を更新完了: " + todayFormatted);
}

function appendDataToMain(newData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MAIN);
  
  if (!mainSheet) {
    throw new Error(`${CONFIG.SHEET_NAMES.MAIN} シートが見つかりません。`);
  }
  
  if (!newData || newData.length === 0) {
    if (typeof writeLog === 'function') writeLog("⚠️ 追記するデータがありませんでした。", "warn");
    return;
  }
  
  const lastRow = mainSheet.getLastRow();
  const startRow = lastRow + 1;
  const numRows = newData.length;
  const numCols = newData[0].length;
  
  mainSheet.getRange(startRow, 1, numRows, numCols).setValues(newData);
  
  if (typeof writeLog === 'function') writeLog(`✅ データベースに ${numRows} 件のデータを新規追記しました。`, "success");
}