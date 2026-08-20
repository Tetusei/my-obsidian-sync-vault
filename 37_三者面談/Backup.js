/**
 * 三者面談 予約システム — バックアップ管理
 * 
 * 現在のスプレッドシートを「📦_バックアップ保存箱」フォルダに日時付きで複製保存する。
 */

/**
 * 手動バックアップを実行する
 */
function createManualBackup() {
  var ssFile = DriveApp.getFileById(SPREADSHEET_ID);
  var folder = getOrCreateBackupFolder_();
  var dateStr = Utilities.formatDate(new Date(), TZ, 'yyyyMMdd_HHmm');
  var backupName = ssFile.getName() + '_バックアップ_' + dateStr;

  var backupFile = ssFile.makeCopy(backupName, folder);

  return {
    fileName: backupFile.getName(),
    fileUrl: backupFile.getUrl(),
    folderName: BACKUP_FOLDER_NAME,
    folderUrl: folder.getUrl()
  };
}

/**
 * バックアップ保存先フォルダを取得または自動生成
 */
function getOrCreateBackupFolder_() {
  var ssFile = DriveApp.getFileById(SPREADSHEET_ID);
  var parents = ssFile.getParents();
  var parentFolder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();

  var folders = parentFolder.getFoldersByName(BACKUP_FOLDER_NAME);
  if (folders.hasNext()) {
    return folders.next();
  }
  return parentFolder.createFolder(BACKUP_FOLDER_NAME);
}

/**
 * バックアップ完了モーダルダイアログ
 */
function showBackupCompleteDialog_(res) {
  var html = '<!DOCTYPE html><html><head><base target="_blank">' +
    '<style>' +
    'body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; padding: 16px 20px; color: #202124; line-height: 1.5; margin: 0; }' +
    'h2 { font-size: 1.15rem; margin: 0 0 12px; color: #137333; display: flex; align-items: center; gap: 8px; }' +
    'p { margin: 0 0 10px; font-size: 0.95rem; }' +
    '.filename { font-size: 0.88rem; color: #5f6368; background: #f1f3f4; padding: 6px 10px; border-radius: 6px; word-break: break-all; margin: 10px 0 16px; }' +
    '.btn-group { display: flex; flex-direction: column; gap: 10px; margin-top: 18px; }' +
    'a.btn { display: block; text-align: center; padding: 10px 14px; text-decoration: none; border-radius: 6px; font-weight: 600; font-size: 0.95rem; box-sizing: border-box; }' +
    'a.btn-primary { background: #137333; color: #ffffff; }' +
    'a.btn-primary:hover { background: #0d5424; }' +
    'a.btn-secondary { background: #ffffff; color: #137333; border: 1px solid #dadce0; }' +
    'a.btn-secondary:hover { background: #f8f9fa; }' +
    'button.btn-close { width: 100%; padding: 8px; background: transparent; border: none; color: #5f6368; font-size: 0.88rem; cursor: pointer; margin-top: 6px; }' +
    'button.btn-close:hover { text-decoration: underline; }' +
    '</style></head><body>' +
    '<h2>📦 バックアップを作成しました</h2>' +
    '<p>現在のスプレッドシートの完全な複製を<strong>「' + res.folderName + '」</strong>に保存しました。</p>' +
    '<div class="filename">📎 ' + res.fileName + '</div>' +
    '<div class="btn-group">' +
    '<a class="btn btn-primary" href="' + res.folderUrl + '">📁 バックアップ保存箱を開く</a>' +
    '<a class="btn btn-secondary" href="' + res.fileUrl + '">📄 バックアップファイルを開く</a>' +
    '<button class="btn-close" onclick="google.script.host.close()">閉じる</button>' +
    '</div></body></html>';

  var userInterface = HtmlService.createHtmlOutput(html)
    .setWidth(420)
    .setHeight(270);
  SpreadsheetApp.getUi().showModalDialog(userInterface, 'バックアップ完了');
}
