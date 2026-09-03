/**
 * 三者面談 予約システム — バックアップ管理
 * 
 * 現在のスプレッドシートを「📦_バックアップ保存箱」フォルダに日時付きで複製保存する。
 * 毎日の自動バックアップは、その中の「🕒_自動（毎日）」サブフォルダへ分けて置く。
 */

/**
 * バックアップを1つ作る。
 * @param {boolean} auto 毎日の自動バックアップなら true
 */
function createBackup_(auto) {
  var ssFile = DriveApp.getFileById(ssId_());
  var folder = auto ? getOrCreateAutoBackupFolder_() : getOrCreateBackupFolder_();
  var dateStr = Utilities.formatDate(new Date(), TZ, 'yyyyMMdd_HHmm');
  var backupName = ssFile.getName() + (auto ? '_自動バックアップ_' : '_バックアップ_') + dateStr;

  var backupFile = ssFile.makeCopy(backupName, folder);

  return {
    fileName: backupFile.getName(),
    fileUrl: backupFile.getUrl(),
    folderName: auto ? BACKUP_FOLDER_NAME + ' ▸ ' + AUTO_BACKUP_FOLDER_NAME : BACKUP_FOLDER_NAME,
    folderUrl: folder.getUrl()
  };
}

/** 手動バックアップを実行する */
function createManualBackup() {
  return createBackup_(false);
}

/** 毎日の自動バックアップを実行する */
function createAutoBackup() {
  return createBackup_(true);
}

/**
 * バックアップ保存先フォルダを取得または自動生成
 */
function getOrCreateBackupFolder_() {
  var ssFile = DriveApp.getFileById(ssId_());
  var parents = ssFile.getParents();
  var parentFolder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();

  var folders = parentFolder.getFoldersByName(BACKUP_FOLDER_NAME);
  if (folders.hasNext()) {
    return folders.next();
  }
  return parentFolder.createFolder(BACKUP_FOLDER_NAME);
}

/**
 * 自動バックアップの保存先。保存箱の中のサブフォルダ。
 * まとめて整理するときに、手動で取った分を巻き込まずに済む。
 */
function getOrCreateAutoBackupFolder_() {
  var parent = getOrCreateBackupFolder_();
  var folders = parent.getFoldersByName(AUTO_BACKUP_FOLDER_NAME);
  if (folders.hasNext()) {
    return folders.next();
  }
  return parent.createFolder(AUTO_BACKUP_FOLDER_NAME);
}

/**
 * 自動バックアップを新しい順に keep 個だけ残し、古いものはゴミ箱へ移す。
 *
 * 完全削除ではなくゴミ箱へ入れるので、間違いに気づいたら Drive のゴミ箱から戻せる。
 * 自動バックアップの名前で始まるファイルだけを対象にし、
 * 人が同じフォルダに置いた別のファイルには触らない。
 * 手動バックアップは別フォルダにあるので、そもそも対象外。
 *
 * @param {number} keep 残す世代数。0以下なら何もしない（無制限）
 * @return {number} ゴミ箱へ移した数
 */
function pruneAutoBackups_(keep) {
  if (!(keep > 0)) return 0;

  var prefix = DriveApp.getFileById(ssId_()).getName() + '_自動バックアップ_';
  var folder = getOrCreateAutoBackupFolder_();

  var files = [];
  var it = folder.getFiles();
  while (it.hasNext()) {
    var f = it.next();
    if (f.getName().indexOf(prefix) !== 0) continue;
    // 並べ替えのたびに API を呼ばないよう、作成日時はここで取っておく
    files.push({ file: f, at: f.getDateCreated().getTime() });
  }
  if (files.length <= keep) return 0;

  files.sort(function (a, b) { return b.at - a.at; });

  var trashed = 0;
  for (var i = keep; i < files.length; i++) {
    try {
      files[i].file.setTrashed(true);
      trashed++;
    } catch (e) {
      console.warn('古い自動バックアップを片付けられませんでした: ' + files[i].file.getName() + ' / ' + e);
    }
  }
  return trashed;
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
