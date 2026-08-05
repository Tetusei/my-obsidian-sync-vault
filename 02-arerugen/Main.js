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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MAIN); 
  const baseSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.CONFIG); 
  
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

  // Mainシートから最新のデータを取得
  if (!mainSheet) return;
  let mainData = mainSheet.getDataRange().getValues();
  if (mainData.length <= 2) {
    clearDashboardList_(dashSheet);
    return;
  }
  
  mainData = mainData.slice(CONFIG.START_ROW - 1); 

  let excludeKeywords = [];
  if (baseSheet) {
    const excludeValue = baseSheet.getRange("B10").getValue();
    if (excludeValue) {
      excludeKeywords = String(excludeValue).split(/[、，,\s\n\r]+/).map(k => k.trim()).filter(Boolean);
    }
  }

  const todaysData = mainData.filter(row => {
    const dateCell = row[1]; 
    if (!dateCell) return false;
    
    const dateStr = String(dateCell).trim();
    let is毎日 = false;
    let rowDateStr = "";
    
    if (dateStr === "毎日") {
      is毎日 = true; 
    } else if (dateCell instanceof Date) {
      rowDateStr = Utilities.formatDate(dateCell, "JST", "yyyy/MM/dd");
    } else {
      const m = dateStr.match(/^(\d{4})[\/\-年](\d{1,2})[\/\-月](\d{1,2})/);
      if (m) {
        rowDateStr = m[1] + "/" + ("0" + m[2]).slice(-2) + "/" + ("0" + m[3]).slice(-2);
      }
    }
    
    if (!is毎日 && rowDateStr !== targetDateStr) return false;
    
    if (is毎日 && excludeKeywords.length > 0) {
      const dishName = String(row[5] || "").trim(); 
      const isExcluded = excludeKeywords.some(keyword => dishName.includes(keyword));
      if (isExcluded) return false; 
    }
    
    return true;
  });

  // --- 集約処理（同じ名前・クラスの生徒を1つに合体） ---
  const groupedMap = {};
  todaysData.forEach(row => {
    const className = String(row[2] || "").trim();    
    const studentName = String(row[3] || "").trim();  
    const groupKey = className + "_" + studentName;

    const dishName = String(row[5] || "").trim();     
    const actionText = String(row[6] || "").trim();   
    
    const formattedAction = dishName ? `【${dishName}】${actionText}` : actionText;

    if (!groupedMap[groupKey]) {
      groupedMap[groupKey] = {
        className: className,
        studentName: studentName,
        allergen: String(row[4] || ""),               
        actions: formattedAction ? [formattedAction] : []
      };
    } else {
      const current = groupedMap[groupKey];
      if (formattedAction && !current.actions.includes(formattedAction)) {
        current.actions.push(formattedAction); 
      }
      if (row[4]) {
        const currentAllergens = current.allergen.split(/[、,，\s]+/).map(a => a.trim()).filter(Boolean);
        const newAllergens = String(row[4]).split(/[、,，\s]+/).map(a => a.trim()).filter(Boolean);
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
}

function clearDashboardList_(dashSheet) {
  // 🌟B〜E列の文字だけを消す（F列・G列には一切触れない）
  const range = dashSheet.getRange("B8:E30");
  range.clearContent();
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