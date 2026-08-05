/**
 * ファイル名: OnEdit.gs
 * 役割: セルの編集イベント検知（APIキー自動保存・非表示化）
 * バージョン: v3.1.0 (シンプルトリガー完全統合・不要コード削除)
 */

/**
 * シンプルトリガーとしてのonEdit
 * ※DocumentPropertiesはシンプルトリガーからでも書き込みが許可されているため、
 * 特別なメニュー操作や事前設定なしで、入力してエンターキーを押すだけで自動で「設定済み」にマスクされます。
 */
function onEdit(e) {
  try {
    processApiKeyEdit(e);
  } catch (err) {
    console.warn("Simple onEdit failed: " + err.message);
  }
}

/**
 * APIキーの退避と「設定済み」マスク表示処理の共通ロジック
 */
function processApiKeyEdit(e) {
  if (!e) return;
  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();
  
  // 「基礎データ」シート以外は無視（Config.jsの定義を使わず、文字列で直接判定してロードエラーを防ぎます）
  if (sheetName.trim() !== "基礎データ") return;
  
  const row = range.getRow();
  const col = range.getColumn();
  const lastRow = range.getLastRow();
  const lastCol = range.getLastCol();
  
  // APIキーの入力範囲（B1:B3）と交差しているか判定
  // B1:B3 は行1〜3、列2（B列）
  if (col <= 2 && lastCol >= 2 && row <= 3 && lastRow >= 1) {
    const docProperties = PropertiesService.getDocumentProperties();
    let changed = false;
    
    for (let r = 1; r <= 3; r++) {
      if (r >= row && r <= lastRow) {
        const cell = sheet.getRange(r, 2);
        const val = cell.getValue();
        const valStr = val ? val.toString().trim() : "";
        const propKey = 'API_KEY_' + (r - 1);
        
        if (valStr === "") {
          // 空欄にされた場合はプロパティから削除
          docProperties.deleteProperty(propKey);
        } else if (valStr === "設定済み") {
          // すでに設定済みの表示なので何もしない
        } else {
          // 新しいAPIキーが入力された場合
          docProperties.setProperty(propKey, valStr);
          cell.setValue("設定済み");
          changed = true;
        }
      }
    }
    
    if (changed) {
      SpreadsheetApp.flush();
    }
  }
}
