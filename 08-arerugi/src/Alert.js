/**
 * Alert.gs (v4.5.7)
 * 役割：「個人アレルゲン」と「献立マスタ」を照合し、保護者からの「Main」シートへの記入漏れを自動検知する
 * 特徴：C列原材料ダイレクト部分一致判定。
 * 　　　基礎データB10セルに記載されたアレルゲン（飲用牛乳など）を完全スルーする機能。
 * 　　　危険と判定された「加工品のブロック」だけを自動抽出してアラートに表示する機能
 * 　　　抽出したテキスト内の「原因キーワード」だけを赤字＋太字でハイライト表示する機能
 * 　　　【強化】献立マスタ内の「【牛乳】」ブロック（飲用牛乳）を完全にアラート対象から抹殺する機能
 */

function runMissingEntryCheck() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast("C列の原材料から記入漏れを正確に検査しています...", "⏳ 検査中");

  // 1. 各シートの取得と存在確認
  const personalSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.PERSONAL);
  const masterSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MASTER);
  const mainSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MAIN);
  const dictSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.DICT);

  if (!personalSheet || !masterSheet || !mainSheet || !dictSheet) {
    SpreadsheetApp.getUi().alert("❌ 必要なシート（個人アレルゲン、献立マスタ、Main、アレルゲン辞書）のいずれかが見つかりません。");
    return;
  }

  // 🌟 基礎データB10セルから「判定無視ワード（例：飲用牛乳）」を取得
  const ignoreWords = typeof CONFIG.getIgnoreWords === 'function' ? CONFIG.getIgnoreWords() : [];

  // 2. アレルゲン辞書の読み込み
  const dictData = dictSheet.getDataRange().getValues();
  const dict = [];
  for (let i = 1; i < dictData.length; i++) {
    if (!dictData[i][0]) continue;
    const category = String(dictData[i][0]).trim();
    const keywords = String(dictData[i][1]).split(/[,、，]/).map(k => k.trim()).filter(k => k);
    const exclusions = String(dictData[i][2] || "").split(/[,、，]/).map(k => k.trim()).filter(k => k);
    dict.push({ category, keywords, exclusions });
  }

  // 3. Mainシート（保護者からの提出済みデータ）の読み込み
  const mainData = mainSheet.getDataRange().getValues();
  const submittedMap = new Map(); 
  
  for (let i = 2; i < mainData.length; i++) {
    const row = mainData[i];
    if (!row[3]) continue; 
    
    const dateStr = (String(row[1]).trim() === "毎日") ? "毎日" : Utilities.formatDate(new Date(row[1]), "JST", "yyyy/MM/dd");
    const name = String(row[3]).normalize("NFKC").replace(/[\s　]/g, "");
    const key = `${dateStr}_${name}`;
    submittedMap.set(key, true);
  }

  // 4. 個人アレルゲン（ベースとなる名簿）の読み込み
  const personalData = personalSheet.getDataRange().getValues();
  const students = [];
  for (let i = 1; i < personalData.length; i++) {
    const cName = String(personalData[i][0] || "").trim();
    const sName = String(personalData[i][1] || "").trim();
    const allergens = String(personalData[i][2] || "").trim();
    if (cName && sName && allergens) {
      students.push({ 
        className: cName, 
        name: sName, 
        cleanName: sName.normalize("NFKC").replace(/[\s　]/g, ""), 
        allergens: allergens 
      });
    }
  }

  // 5. 献立マスタ（C列のみ）との照合チェック
  const masterData = masterSheet.getDataRange().getValues();
  const alerts = [];

  for (let i = CONFIG.START_ROW - 1; i < masterData.length; i++) {
    const row = masterData[i];
    if (!row[0]) continue;

    const mDateObj = new Date(row[0]);
    if (isNaN(mDateObj.getTime())) continue;
    const dateStr = Utilities.formatDate(mDateObj, "JST", "yyyy/MM/dd");
    const menuName = String(row[1] || "");
    const ingredients = String(row[2] || "");

    students.forEach(student => {
      let isDangerous = false;
      let foundKeywords = []; 

      // 児童のアレルゲン申告を単語に分解
      const allergenWords = student.allergens.split(/[・、，,\s\n\r]+/).map(a => a.trim()).filter(Boolean);

      allergenWords.forEach(word => {
        if (word === "飲用牛乳") return; // 単語自体が飲用牛乳ならスキップ
        if (ignoreWords.includes(word)) return; 

        dict.forEach(d => {
          let isExcludedAllergen = false;
          d.exclusions.forEach(ex => {
            if (ex && word.includes(ex)) isExcludedAllergen = true; 
          });

          if (!isExcludedAllergen) {
            let isMatch = false;
            
            if (word === d.category || word.includes(d.category) || d.category.includes(word)) {
              isMatch = true;
            }
            if (word.includes("牛肉") && d.category === "肉") {
              isMatch = false; 
            }

            if (isMatch) {
              d.keywords.forEach(kw => {
                if (ingredients.includes(kw)) {
                  let isExcludedWord = false;
                  
                  if (kw === "牛" && (ingredients.includes("牛乳") || ingredients.includes("牛蒡") || ingredients.includes("ごぼう"))) {
                    const testString = ingredients.split("牛乳").join("").split("牛蒡").join("").split("ごぼう").join("");
                    if (!testString.includes("牛")) isExcludedWord = true; 
                  }
                  
                  if (kw === "肉" && (ingredients.includes("魚肉") || ingredients.includes("梅肉"))) {
                    const testString = ingredients.split("魚肉").join("").split("梅肉").join("");
                    if (!testString.includes("肉")) isExcludedWord = true; 
                  }

                  let isExcludedByDict = false;
                  d.exclusions.forEach(ex => {
                    if (ex && ingredients.includes(ex)) {
                      if (kw === ex) isExcludedByDict = true;
                    }
                  });

                  if (!isExcludedWord && !isExcludedByDict) {
                    isDangerous = true;
                    foundKeywords.push(kw); 
                  }
                }
              });
            }
          }
        });
      }); 

      // ⚠️ 危険と判定された場合の処理
      if (isDangerous && foundKeywords.length > 0) {
        const mainKey = `${dateStr}_${student.cleanName}`;
        const dailyKey = `毎日_${student.cleanName}`;
        
        let status = "❌ 未提出（完全な記入漏れの疑い）";
        if (submittedMap.has(mainKey) || submittedMap.has(dailyKey)) {
           status = "⚠️ 提出あり（対象料理が正しく記入されているか要確認）";
        }

        // C列を【 で区切ってブロックに分け、キーワードが含まれるものだけを抽出
        const matchedBlocks = [];
        const blocks = ingredients.split(/(?=【)/); 
        
        blocks.forEach(block => {
          const hasKeyword = foundKeywords.some(kw => block.includes(kw));
          if (hasKeyword && block.trim() !== "") {
            
            // 🛑【ここを強化！】もしブロックが「【牛乳】」だったら、最初から配らない物なのでアラートから完全除外
            if (block.includes("【牛乳】")) {
              return; 
            }
            
            matchedBlocks.push(block.trim());
          }
        });

        // 🛑【ここを追加！】除外した結果、危険なブロックが1つも残らなかったら、アラート自体を出さない！
        if (matchedBlocks.length === 0) {
          return; 
        }

        const uniqueKeywordsArray = [...new Set(foundKeywords)];
        const uniqueKeywords = uniqueKeywordsArray.join("、");
        const uniqueBlocks = [...new Set(matchedBlocks)].join("\n");
        const alertContent = `[原因: ${uniqueKeywords}]\n${uniqueBlocks}`;

        alerts.push([
          dateStr,
          menuName,
          student.className,
          student.name,
          student.allergens,
          alertContent,
          status,
          false,
          uniqueKeywordsArray 
        ]);
      }
    });
  }

  // 🌟🌟 出力と赤字ハイライト処理 🌟🌟
  let alertSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.ALERT);
  if (!alertSheet) {
    alertSheet = ss.insertSheet(CONFIG.SHEET_NAMES.ALERT);
  }
  alertSheet.clear();

  const headers = [["日付", "献立名", "クラス", "氏名", "アレルゲン(個人票)", "検出された危険素材", "Mainシート提出状況", "先生の確認"]];
  alertSheet.getRange(1, 1, 1, headers[0].length).setValues(headers)
            .setBackground("#d93025").setFontColor("white").setFontWeight("bold");

  if (alerts.length > 0) {
    alerts.sort((a, b) => {
      if (a[0] !== b[0]) return a[0] > b[0] ? 1 : -1;
      if (a[2] !== b[2]) return a[2] > b[2] ? 1 : -1;
      return a[3] > b[3] ? 1 : -1;
    });

    const outputData = alerts.map(row => row.slice(0, 8));
    const dataRange = alertSheet.getRange(2, 1, outputData.length, 8);
    dataRange.setValues(outputData);
    dataRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP); 
    dataRange.setVerticalAlignment("middle");

    alertSheet.getRange(2, 8, outputData.length, 1).insertCheckboxes();

    // 🎨 F列の危険キーワードを赤字＋太字にする処理
    const redStyle = SpreadsheetApp.newTextStyle().setForegroundColor("red").setBold(true).build();
    const richTextValues = alerts.map(row => {
      const text = String(row[5]);
      const keywords = row[8]; 
      const builder = SpreadsheetApp.newRichTextValue().setText(text);

      keywords.forEach(kw => {
        if (!kw) return;
        let startIndex = 0;
        while ((startIndex = text.indexOf(kw, startIndex)) !== -1) {
          builder.setTextStyle(startIndex, startIndex + kw.length, redStyle);
          startIndex += kw.length;
        }
      });
      return [builder.build()];
    });
    alertSheet.getRange(2, 6, richTextValues.length, 1).setRichTextValues(richTextValues);

    alertSheet.setColumnWidth(1, 90);
    alertSheet.setColumnWidth(2, 180);
    alertSheet.setColumnWidth(3, 80);
    alertSheet.setColumnWidth(4, 100);
    alertSheet.setColumnWidth(5, 150);
    alertSheet.setColumnWidth(6, 320); 
    alertSheet.setColumnWidth(7, 280);

    const bgColors = [];
    let isColored = false;
    let prevDateStr = "";

    for (let i = 0; i < outputData.length; i++) {
      const currentDateStr = outputData[i][0];
      if (i > 0 && currentDateStr !== prevDateStr) isColored = !isColored;
      prevDateStr = currentDateStr;

      const baseColor = isColored ? "#e8f0fe" : "#ffffff";
      const rowColors = Array(8).fill(baseColor);

      if (outputData[i][6].includes("未提出")) {
        rowColors[6] = "#fce8e6"; 
      } else {
        rowColors[6] = "#fef7e0"; 
      }
      bgColors.push(rowColors);
    }
    dataRange.setBackgrounds(bgColors);

  } else {
    alertSheet.getRange(2, 1).setValue("🎉 記入漏れや危険な項目は見つかりませんでした！");
  }

  ss.setActiveSheet(alertSheet);
  ss.toast("チェックが完了しました。アラートシートを確認してください。", "✅ 完了", 8);
  if (typeof writeLog === 'function') writeLog(`🚨 記入漏れチェック完了: ${alerts.length}件の警告を検出`, "info");
}