/**
 * Verify.gs
 * 役割：裏DB蓄積型・月別抽出プログラム
 * （同一人・複数対応1行集約 ＋ 毎日展開 ＋ パン粉誤爆防止 ＋ 幽霊データ一掃 完全版）
 */

// 1. 一覧作成のメイン処理
function createMonthlyVerifyList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const verifySheet = ss.getSheetByName("確認用一覧");
  const mainSheet = ss.getSheetByName("Main");
  const masterSheet = ss.getSheetByName("献立マスタ");
  const dbSheet = ss.getSheetByName("AllDB");
  
  if (typeof writeLog === 'function') writeLog("🔄 一覧作成処理を開始します...", "info");
  
  const targetMonthLabel = verifySheet.getRange("C1").getValue(); 
  
  try {
    const latestDbData = saveToAllDB(mainSheet, masterSheet, dbSheet);
    updateVerifyDisplay(latestDbData, verifySheet, targetMonthLabel);
    
    SpreadsheetApp.flush(); 
    
    ss.toast(`✅ ${targetMonthLabel} の一覧を表示しました`, "完了", 8);
    if (typeof writeLog === 'function') writeLog(`✅ 一覧作成完了: データを画面に反映しました`, "success");
    
  } catch (error) {
    ss.toast("❌ エラーが発生しました: " + error.message, "処理中断", 10);
    if (typeof writeLog === 'function') writeLog(`❌ 一覧作成エラー: ${error.message}`, "error");
  }
}

// 2. データベースへの保存・上書き防止・毎日展開処理
function saveToAllDB(mainSheet, masterSheet, dbSheet) {
  const mainData = mainSheet.getDataRange().getValues();
  const masterLastRow = masterSheet.getLastRow();
  const masterData = masterLastRow > 1 ? masterSheet.getRange("A2:D" + masterLastRow).getValues() : [];
  
  if (mainData.length <= 1) return [];

  const masterDict = {};
  masterData.forEach(row => {
    if (row[0]) {
      const d = new Date(row[0]);
      if (!isNaN(d.getTime())) {
        const dateKey = Utilities.formatDate(d, "JST", "yyyy/MM/dd");
        masterDict[dateKey] = { name: row[1], ingredients: row[2], memo: row[3] };
      }
    }
  });

  let dbData = dbSheet.getDataRange().getValues();
  const dbHeader = dbData.length > 0 ? dbData.shift() : ["日付","献立名","材料","クラス","氏名","アレルゲン","対応内容","辞書登録"];
  
  const createKey = (date, name, dish) => {
    const d = new Date(date);
    if (isNaN(d.getTime())) return "InvalidDate_" + Math.random();
    const cleanNameForKey = String(name || "").normalize("NFKC").replace(/[\s　]/g, "");
    const cleanDishForKey = String(dish || "").normalize("NFKC").replace(/[\s　]/g, ""); 
    return Utilities.formatDate(d, "JST", "yyyy/MM/dd") + "_" + cleanNameForKey + "_" + cleanDishForKey;
  };

  const dbMap = {};
  dbData.forEach(row => {
    if (row[0]) {
      const actionText = row[6] ? String(row[6]) : "";
      const matchDish = actionText.match(/^【(.*?)】/);
      const dishKey = matchDish ? matchDish[1] : "";
      dbMap[createKey(row[0], row[4], dishKey)] = row;
    }
  });

  mainData.forEach((row, i) => {
    if (i < 2 || !row[1]) return; 
    
    const dishName = row[5] ? String(row[5]).trim() : "";
    const cleanNameForKey = String(row[3] || "").normalize("NFKC").replace(/[\s　]/g, "");
    const cleanDishForKey = dishName.normalize("NFKC").replace(/[\s　]/g, "");

    // 🌟「毎日」指定の場合の特別展開処理
    if (String(row[1]).trim() === "毎日") {
      
      // 🌟【重要】まず、AllDB（過去の記録）からこの児童のこの料理データを全日程分一旦「完全に削除」する
      // （以前のプログラムで誤爆して保存されてしまった「幽霊データ」を消し去るため）
      Object.keys(dbMap).forEach(key => {
        const parts = key.split("_");
        if (parts.length >= 3) {
          const keyName = parts[1];
          const keyDish = parts.slice(2).join("_");
          if (keyName === cleanNameForKey && keyDish === cleanDishForKey) {
            delete dbMap[key];
          }
        }
      });

      // チェックが入っている場合のみ、スマートフィルターを通して再登録する
      if (row[0] === true) {
        Object.keys(masterDict).forEach(dateKeyStr => {
          const dateParts = dateKeyStr.split("/");
          const dateObj = new Date(dateParts[0], dateParts[1] - 1, dateParts[2]);
          const uniqueKey = dateKeyStr + "_" + cleanNameForKey + "_" + cleanDishForKey;
          
          const menuInfo = masterDict[dateKeyStr] || { name: "（未登録）", ingredients: "（未登録）", memo: "" };
          
          // スマートフィルター（パン粉などの誤爆防止）
          if (dishName) {
            const menuText = String(menuInfo.name + " " + menuInfo.ingredients);
            if (dishName.includes("パン")) {
              const textWithoutPanko = menuText.replace(/パン粉/g, "");
              if (!textWithoutPanko.includes("パン")) {
                return; // パンがない日はスキップ（AllDBにも追加されない）
              }
            }
            if (dishName.includes("牛乳") && !menuText.includes("牛乳") && !menuText.includes("ミルク") && !menuText.includes("乳")) {
              return; 
            }
          }

          let actionText = row[6] ? String(row[6]).trim() : "";
          if (dishName) {
            actionText = actionText.replace(/^【.*?】\s*/, '');
            actionText = `【${dishName}】${actionText}`;
          }

          const cleanClass = String(row[2] || "").replace(/[\s　]/g, "").normalize("NFKC");
          const cleanName = String(row[3] || "").normalize("NFKC");

          const newRow = [
            dateObj, 
            menuInfo.name, 
            menuInfo.ingredients, 
            cleanClass, 
            cleanName,  
            row[4],     
            actionText, 
            menuInfo.memo 
          ];
          dbMap[uniqueKey] = newRow;
        });
      }
      return; 
    }
    
    // 🌟通常の日付処理
    const dateObj = new Date(row[1]);
    if (isNaN(dateObj.getTime()) || dateObj.getFullYear() < 2000) return;
    
    const uniqueKey = createKey(dateObj, row[3], dishName);

    if (row[0] !== true) {
      if (dbMap[uniqueKey]) {
        delete dbMap[uniqueKey];
      }
      return; 
    }
    
    const dateKeyStr = Utilities.formatDate(dateObj, "JST", "yyyy/MM/dd");
    const menuInfo = masterDict[dateKeyStr] || { name: "（未登録）", ingredients: "（未登録）", memo: "" };
    
    let actionText = row[6] ? String(row[6]).trim() : "";
    
    if (dishName) {
      actionText = actionText.replace(/^【.*?】\s*/, '');
      actionText = `【${dishName}】${actionText}`;
    }

    const cleanClass = String(row[2] || "").replace(/[\s　]/g, "").normalize("NFKC");
    const cleanName = String(row[3] || "").normalize("NFKC");

    const newRow = [
      dateObj, 
      menuInfo.name, 
      menuInfo.ingredients, 
      cleanClass, 
      cleanName,  
      row[4],     
      actionText, 
      menuInfo.memo 
    ];
    dbMap[uniqueKey] = newRow;
  });

  const updatedDbData = Object.keys(dbMap).map(key => dbMap[key]);
  updatedDbData.sort((a, b) => {
    const t1 = new Date(a[0]).getTime();
    const t2 = new Date(b[0]).getTime();
    if (t1 !== t2) return t1 - t2;
    return String(a[3]).localeCompare(String(b[3])); 
  });

  dbSheet.clearContents();
  dbSheet.getRange(1, 1, 1, dbHeader.length).setValues([dbHeader]);
  if (updatedDbData.length > 0) {
    dbSheet.getRange(2, 1, updatedDbData.length, dbHeader.length).setValues(updatedDbData);
    dbSheet.getRange(2, 1, updatedDbData.length, 1).setNumberFormat("yyyy/MM/dd");
  }

  const fullRange = dbSheet.getRange(1, 1, updatedDbData.length + 1, dbHeader.length);
  fullRange.setFontColor("#000000"); 

  if (typeof writeLog === 'function') {
    writeLog(`✅ データベースに ${updatedDbData.length} 件のデータを保存・更新しました。`, "success");
  }
  
  return [dbHeader, ...updatedDbData]; 
}

// 3. 確認用一覧シートへの表示と色付け処理
function updateVerifyDisplay(allData, verifySheet, targetMonthLabel) {
  if (!Array.isArray(allData) || allData.length === 0) return;
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const baseSheet = ss.getSheetByName("基礎データ");
  let excludeKeywords = [];
  if (baseSheet) {
    const excludeValue = baseSheet.getRange("B10").getValue();
    if (excludeValue) {
      excludeKeywords = String(excludeValue).split(/[、，,\s\n\r]+/).map(k => k.trim()).filter(Boolean);
    }
  }
  
  const dataCopy = allData.slice();
  const header = dataCopy.shift(); 
  
  const filtered = dataCopy.filter(row => {
    if (targetMonthLabel !== "全期間" && targetMonthLabel !== "") {
      if (((new Date(row[0]).getMonth() + 1) + "月") !== targetMonthLabel) return false;
    }

    const actionText = String(row[6] || ""); 
    const matchDish = actionText.match(/^【(.*?)】/);
    const dishName = matchDish ? matchDish[1].trim() : "";
    
    if (dishName && excludeKeywords.length > 0) {
      const isExcluded = excludeKeywords.some(keyword => dishName.includes(keyword));
      if (isExcluded) return false; 
    }
    return true;
  });

  const groupedMap = {};
  filtered.forEach(row => {
    const d = new Date(row[0]);
    if (isNaN(d.getTime())) return;
    const dateKey = Utilities.formatDate(d, "JST", "yyyy/MM/dd");
    const className = String(row[3] || "").trim();
    const studentName = String(row[4] || "").trim();
    const groupKey = dateKey + "_" + className + "_" + studentName;

    if (!groupedMap[groupKey]) {
      groupedMap[groupKey] = {
        date: row[0],
        menuName: String(row[1] || ""),
        ingredients: String(row[2] || ""),
        className: row[3],
        studentName: row[4],
        allergen: String(row[5] || ""),
        actions: [String(row[6] || "")],
        memo: String(row[7] || "")
      };
    } else {
      const current = groupedMap[groupKey];
      
      const actionStr = String(row[6] || "");
      if (actionStr && !current.actions.includes(actionStr)) {
        current.actions.push(actionStr);
      }
      
      if (row[5]) {
        const currentAllergens = current.allergen.split(/[、,，\s]+/).map(a => a.trim()).filter(Boolean);
        const newAllergens = String(row[5]).split(/[、,，\s]+/).map(a => a.trim()).filter(Boolean);
        current.allergen = [...new Set([...currentAllergens, ...newAllergens])].join("、");
      }
      
      if (row[7]) {
        const currentMemos = current.memo.split("\n").map(m => m.trim()).filter(Boolean);
        const newMemos = String(row[7]).split("\n").map(m => m.trim()).filter(Boolean);
        current.memo = [...new Set([...currentMemos, ...newMemos])].join("\n");
      }
    }
  });

  const groupedList = Object.keys(groupedMap).map(key => {
    const g = groupedMap[key];
    return [
      g.date,
      g.menuName,
      g.ingredients,
      g.className,
      g.studentName,
      g.allergen,
      g.actions.join("\n"),
      g.memo
    ];
  });

  groupedList.sort((a, b) => {
    const dA = new Date(a[0]);
    const dB = new Date(b[0]);
    const dateA = isNaN(dA.getTime()) ? "99999999" : Utilities.formatDate(dA, "JST", "yyyyMMdd");
    const dateB = isNaN(dB.getTime()) ? "99999999" : Utilities.formatDate(dB, "JST", "yyyyMMdd");
    if (dateA !== dateB) return dateA < dateB ? -1 : 1; 
    
    const classA = String(a[3] || "");
    const classB = String(b[3] || "");
    if (classA !== classB) return classA < classB ? -1 : 1;

    const nameA = String(a[4] || "");
    const nameB = String(b[4] || "");
    return nameA < nameB ? -1 : 1;
  });
  
  const lastRow = verifySheet.getLastRow();
  const lastCol = verifySheet.getLastColumn() || 8;
  if (lastRow > 2) {
    const clearRange = verifySheet.getRange(3, 1, lastRow - 2, lastCol);
    clearRange.clearContent(); clearRange.setBackground(null); clearRange.setFontColor(null); clearRange.setFontWeight(null); 
  }

  if (groupedList.length > 0) {
    const formattedData = groupedList.map(row => {
      const newRow = [...row]; 
      
      const d = new Date(newRow[0]);
      if (!isNaN(d.getTime())) {
        newRow[0] = Utilities.formatDate(d, "JST", "yyyy/MM/dd") + "(" + ["日", "月", "火", "水", "木", "金", "土"][d.getDay()] + ")";
      }
      
      let ingredientsText = String(newRow[2] || "");
      if (ingredientsText) {
        ingredientsText = ingredientsText.replace(/[【\[].*?[】\]]/g, "");
        const items = ingredientsText.split(/[、，,\s\n\r]+/);
        const uniqueItems = [...new Set(items.map(item => item.trim()).filter(Boolean))];
        newRow[2] = uniqueItems.join("、");
      }
      
      return newRow;
    });
    
    const targetRange = verifySheet.getRange(3, 1, formattedData.length, formattedData[0].length);
    const bgColors = []; const richTextValues = [];
    let isColored = false; let prevDateStr = "";

    const regex = /[【\[](.*?)[】\]]/g;

    for (let i = 0; i < formattedData.length; i++) {
      const currentDateStr = formattedData[i][0];
      if (i > 0 && currentDateStr !== prevDateStr) isColored = !isColored;
      const rowColor = isColored ? "#fff3e0" : "#ffffff";
      
      const colorRow = []; const richTextRow = [];
      const memoText = String(formattedData[i][7] || "");
      const highlights = [];
      let match;
      
      regex.lastIndex = 0; 
      while ((match = regex.exec(memoText)) !== null) {
        const parts = match[1].split(/[：:]/);
        if (parts.length >= 2 && parts[1].trim() !== "") {
          const isAi = parts[0].includes("AI？") || parts[0].includes("AI?");
          highlights.push({ word: parts[1].trim(), color: isAi ? "#f29900" : "#d93025" });
        }
      }

      for (let j = 0; j < formattedData[i].length; j++) {
        colorRow.push(rowColor);
        const cellValue = String(formattedData[i][j] || "");
        let builder = SpreadsheetApp.newRichTextValue().setText(cellValue);
        
        if (j === 2 && highlights.length > 0) {
          highlights.forEach(hl => {
            if (!hl.word || hl.word.length === 0) return; 
            let startIndex = 0, foundIndex;
            while ((foundIndex = cellValue.indexOf(hl.word, startIndex)) !== -1) {
              builder.setTextStyle(foundIndex, foundIndex + hl.word.length, SpreadsheetApp.newTextStyle().setForegroundColor(hl.color).setBold(true).build());
              startIndex = foundIndex + hl.word.length;
            }
          });
        }

        if (j === 7 && cellValue.includes("AI？")) {
          let startIndex = 0, foundIndex;
          while ((foundIndex = cellValue.indexOf("AI？", startIndex)) !== -1) {
            builder.setTextStyle(foundIndex, foundIndex + 3, SpreadsheetApp.newTextStyle().setForegroundColor("#d93025").setBold(true).build());
            startIndex = foundIndex + 3;
          }
        }
        richTextRow.push(builder.build());
      }
      bgColors.push(colorRow); richTextValues.push(richTextRow);
      prevDateStr = currentDateStr;
    }
    targetRange.setBackgrounds(bgColors); targetRange.setRichTextValues(richTextValues);
  }
}