/**
 * Dictionary.gs
 * 役割：辞書の自動学習（編集検知による自動更新）
 */
function processDictionaryLearning(e) {
  const range = e.range;
  const val = String(range.getValue() || "");
  
  // 1. 学習対象（【 】または[ ]）が含まれているかチェック
  const matches = val.match(/[【\[][^：:]+[：:][^】\]]+[】\]]/g);
  if (!matches) return;

  // 2. 「AI？」が含まれていない確定した単語があるかチェック
  const hasConfirmedWord = matches.some(m => !m.includes("AI？") && !m.includes("AI?"));
  if (!hasConfirmedWord) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dictSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.DICT);

  // --- 🌟 バックアップ機能（表示したまま末尾へ移動） ---
  try {
    const now = Utilities.formatDate(new Date(), "JST", "yyyyMMdd_HHmm");
    const backupName = `DictBackup_${now}`;
    if (!ss.getSheetByName(backupName)) {
      const backupSheet = dictSheet.copyTo(ss);
      backupSheet.setName(backupName);
      ss.setActiveSheet(backupSheet);
      ss.moveActiveSheet(ss.getNumSheets());
      ss.setActiveSheet(range.getSheet()); // 元のシートに戻る
      ss.toast(`辞書のバックアップ「${backupName}」を作成しました。`, "💾 Backup");
    }
  } catch (err) {
    console.error("バックアップ作成失敗: " + err);
  }

  // --- 3. 学習ロジック ---
  const lastRow = dictSheet.getLastRow();
  const dictData = lastRow >= CONFIG.START_ROW ? 
      dictSheet.getRange(CONFIG.START_ROW, 1, lastRow - CONFIG.START_ROW + 1, 2).getValues() : [];

  let addedWords = [];

  matches.forEach(m => {
    // AI？がついているものはスルー
    if (m.includes("AI？") || m.includes("AI?")) return;

    const content = m.replace(/[【】\[\]]/g, "");
    const parts = content.split(/[：:]/);
    if (parts.length < 2) return;
    
    // 【重要】ゆらぎ対策：全角半角スペースを削除
    const category = parts[0].trim().replace(/[\s　]/g, ""); 
    const newWord = parts[1].trim().replace(/[\s　]/g, ""); 

    let foundRow = -1;
    let existingWords = "";

    // 既存のカテゴリーがあるかループで確認
    for (let i = 0; i < dictData.length; i++) {
      // 辞書側のカテゴリーもスペースを除去して比較
      const dictCategory = String(dictData[i][0]).trim().replace(/[\s　]/g, "");
      if (dictCategory === category) {
        foundRow = i + CONFIG.START_ROW;
        existingWords = String(dictData[i][1] || "");
        break;
      }
    }

    // 単語の重複チェック
    const wordList = existingWords.split(/[，、,]/).map(w => w.trim());
    if (!wordList.includes(newWord)) {
      if (foundRow !== -1) {
        // --- 既存のカテゴリーに行を追加 ---
        const separator = existingWords ? "，" : "";
        const updatedText = existingWords + separator + newWord;
        
        // 新しい単語の部分だけ「青・太字」にするリッチテキスト作成
        const richText = SpreadsheetApp.newRichTextValue()
          .setText(updatedText)
          .setTextStyle(
            existingWords.length + separator.length, 
            updatedText.length, 
            SpreadsheetApp.newTextStyle().setForegroundColor("#0000FF").setBold(true).build()
          )
          .build();
        
        dictSheet.getRange(foundRow, 2).setRichTextValue(richText);
        // ループ内での連続追加に対応するためデータを更新
        dictData[foundRow - CONFIG.START_ROW][1] = updatedText;
      } else {
        // --- 新しいカテゴリーとして行を追加 ---
        const newRow = [category, newWord];
        dictSheet.appendRow(newRow);
        
        // 追加した行の単語列を「青・太字」にする
        const targetCell = dictSheet.getRange(dictSheet.getLastRow(), 2);
        const richText = SpreadsheetApp.newRichTextValue()
          .setText(newWord)
          .setTextStyle(SpreadsheetApp.newTextStyle().setForegroundColor("#0000FF").setBold(true).build())
          .build();
        targetCell.setRichTextValue(richText);
      }
      addedWords.push(newWord);
    }
  });

  if (addedWords.length > 0) {
    ss.toast(`「${addedWords.join("、")}」を辞書に登録しました。`, "✅ 学習完了");
  }
}