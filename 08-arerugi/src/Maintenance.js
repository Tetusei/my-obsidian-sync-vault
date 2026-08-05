/**
 * Maintenance.gs
 * 役割：データメンテナンス（データの消去など）
 */

/**
 * 運用データ（Main/Master/Verify）を全て一括クリア
 */
function clearAllData() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert("⚠️ 警告", "運用データ（Main/Master/Verify/記入漏れアラート）を全て消去しますか？", ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [CONFIG.SHEET_NAMES.MAIN, CONFIG.SHEET_NAMES.MASTER, CONFIG.SHEET_NAMES.VERIFY].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh && sh.getLastRow() >= CONFIG.START_ROW) {
      sh.getRange(CONFIG.START_ROW, 1, sh.getLastRow() - CONFIG.START_ROW + 1, sh.getLastColumn())
        .clearContent()
        .setBackground(null)
        .removeCheckboxes();
    }
  });

  // 記入漏れアラート（ALERT）シートの2行目以下をクリア
  const alertSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.ALERT);
  if (alertSheet && alertSheet.getLastRow() >= 2) {
    alertSheet.getRange(2, 1, alertSheet.getLastRow() - 1, alertSheet.getLastColumn())
      .clearContent()
      .setBackground(null)
      .removeCheckboxes();
  }

  ss.toast("全ての運用データをクリアしました", "🧹 完了");
}

/**
 * 個人データ(Main)のみクリア
 */
function clearMainDataOnly() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAMES.MAIN);
  if (sh && sh.getLastRow() >= CONFIG.START_ROW) {
    sh.getRange(CONFIG.START_ROW, 1, sh.getLastRow() - CONFIG.START_ROW + 1, sh.getLastColumn())
      .clearContent()
      .setBackground(null)
      .removeCheckboxes();
    SpreadsheetApp.getActiveSpreadsheet().toast("個人データをクリアしました", "🧹 完了");
  }
}

/**
 * 確認用一覧(Verify)のみクリア
 */
function clearVerifyDataOnly() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAMES.VERIFY);
  if (sh && sh.getLastRow() >= CONFIG.START_ROW) {
    // 8列分（A-H列）をクリア
    sh.getRange(CONFIG.START_ROW, 1, sh.getLastRow() - CONFIG.START_ROW + 1, 8)
      .clearContent()
      .setBackground(null);
    SpreadsheetApp.getActiveSpreadsheet().toast("確認用一覧をクリアしました", "🧹 完了");
  }
}

/**
 * 献立マスタ(Master)のみクリア
 */
function clearMasterDataOnly() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert("⚠️ 確認", "献立マスタ（Master）のデータを全て消去しますか？", ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAMES.MASTER);
  if (sh && sh.getLastRow() >= CONFIG.START_ROW) {
    // 3行目（CONFIG.START_ROW）から最終行まで全列クリア
    sh.getRange(CONFIG.START_ROW, 1, sh.getLastRow() - CONFIG.START_ROW + 1, sh.getLastColumn())
      .clearContent()
      .setBackground(null)
      .removeCheckboxes();
    
    SpreadsheetApp.getActiveSpreadsheet().toast("献立マスタをクリアしました", "🧹 完了");
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast("消去するデータがありません", "📝 情報");
  }
}