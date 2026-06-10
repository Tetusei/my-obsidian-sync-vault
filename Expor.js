/**
 * Export.gs
 * 役割：PDF出力、バックアップ作成、引継ぎドキュメント作成
 */

/**
 * 確認用一覧(Verify)をPDFとして保存・表示する
 */
function exportVerifySheetAsPDF() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.VERIFY);
  
  if (!sheet || sheet.getLastRow() < CONFIG.START_ROW) {
    SpreadsheetApp.getUi().alert("出力するデータがありません。");
    return;
  }

  ss.toast("PDFを作成中です...", "⏳ 処理開始");

  const folderName = "📄_PDF帳票保存箱";
  const ssFile = DriveApp.getFileById(ss.getId());
  const parentFolder = ssFile.getParents().hasNext() ? ssFile.getParents().next() : DriveApp.getRootFolder();
  const folders = parentFolder.getFoldersByName(folderName);
  const targetFolder = folders.hasNext() ? folders.next() : parentFolder.createFolder(folderName);

  const sheetId = sheet.getSheetId();
  const url = "https://docs.google.com/spreadsheets/d/" + ss.getId() + "/export" +
              "?format=pdf&size=A4&portrait=false&fitw=true&gridlines=true&gid=" + sheetId;

  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: { 'Authorization': 'Bearer ' + token }
  });

  const dateStr = Utilities.formatDate(new Date(), "JST", "yyyyMMdd_HHmm");
  const fileName = `【確認用一覧】_${dateStr}.pdf`;
  const pdfFile = targetFolder.createFile(response.getBlob()).setName(fileName);

  const htmlText = `
    <div style="font-family: sans-serif; text-align: center;">
      <p>✅ PDFを作成し「${folderName}」へ保存しました。</p>
      <a href="${pdfFile.getUrl()}" target="_blank" 
         style="display: inline-block; background-color: #d93025; color: white; padding: 10px 20px; text-decoration: none; border-radius: 4px; font-weight: bold;">
         📕 PDFを開いて印刷する
      </a>
    </div>
  `;
  const htmlOutput = HtmlService.createHtmlOutput(htmlText).setWidth(400).setHeight(150);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '📄 PDF出力完了');
}

/**
 * スプレッドシート全体のバックアップを作成
 */
function backupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ssFile = DriveApp.getFileById(ss.getId());
  // 🌟 修正：Config.gsの設定値を正しく参照するように統一しました
  const folderName = CONFIG.FOLDERS.BACKUP;
  const parentFolders = ssFile.getParents();
  let parentFolder = parentFolders.hasNext() ? parentFolders.next() : DriveApp.getRootFolder();
  const folders = parentFolder.getFoldersByName(folderName);
  let targetFolder = folders.hasNext() ? folders.next() : parentFolder.createFolder(folderName);

  const dateStr = Utilities.formatDate(new Date(), "JST", "yyyyMMdd_HHmm");
  const backupName = `【${CONFIG.VERSION}_${dateStr}】${ss.getName()}`;
  const backupFile = ssFile.makeCopy(backupName, targetFolder);

  const htmlText = `<div style="font-family: sans-serif; padding: 10px;">
      <p>✅ バックアップを保存しました。</p>
      <p>場所: <strong>${folderName}</strong></p>
      <a href="${backupFile.getUrl()}" target="_blank" style="color: #1a73e8;">📄 バックアップを開く</a>
    </div>`;
  const htmlOutput = HtmlService.createHtmlOutput(htmlText).setWidth(400).setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '💾 バックアップ完了');
  
  if (typeof writeLog === 'function') writeLog(`📦 バックアップ完了: ${backupName}`, "success");
}
