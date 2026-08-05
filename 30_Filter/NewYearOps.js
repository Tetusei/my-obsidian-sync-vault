/**
 * ============================================================================
 * ファイル名: NewYearOps.gs
 * 概要: バックアップ作成、3番目シートの複製、データクリア、色付けを安全に実行する。
 * 排他ロック(LockService)により、実行中は定期自動トリガーを実質OFFにします。
 * ============================================================================
 */

/**
 * 年度更新の一連のシーケンスを安全に実行するメイン関数
 */
function executeNewYearMigrationSequence() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 🔒 1. 自動実行トリガーとの衝突を防ぐため、システム全体をロック（一時的に自動実行をOFF状態にする）
  const lock = LockService.getScriptLock();
  try {
    // 他のトリガーが動いていた場合、最大30秒間待機。取得できなければエラーを出す
    if (!lock.tryLock(30000)) {
      throw new Error('現在、システムの自動実行トリガーが動作中のため、年度更新を開始できません。\nお手数ですが、1〜2分後に再度お試しください。');
    }
    
    ss.toast('1/4: 自動バックアップを作成中...', '処理中', 10000);
    
    // 2. 自動バックアップの作成
    const currentVersion = SystemConfig.VERSION;
    createSystemBackup(ss, currentVersion);
    
    ss.toast('2/4: 新年度シートを作成中...', '処理中', 10000);
    
    // 🎯 3. 【修正】左から3番目（インデックス 2）にある現行の年度シート（例: R8年度）を取得
    // ※ 0:基礎データ, 1:名簿, 2:現行年度シート という並びに対応
    const sheets = ss.getSheets();
    if (sheets.length < 3) {
      throw new Error('シートが3枚未満のため、3番目の年度シートが見つかりません。シートの並び順を確認してください。');
    }
    const currentSheet = sheets[2]; 
    const currentSheetName = currentSheet.getName();
    
    // 4. 次の年度名を自動算出 (例:「R8年度」から「R9年度」を生成)
    const nextSheetName = generateNextSheetName(currentSheetName);
    
    // 既に同名のシートが存在しないか確認（エラー回避）
    if (ss.getSheetByName(nextSheetName)) {
      throw new Error('「' + nextSheetName + '」という名前のシートは既に存在するため、処理を中断しました。');
    }
    
    // 5. 3番目の現行シートをコピーして、左から3番目（インデックス 2）に挿入
    const newSheet = currentSheet.copyTo(ss);
    newSheet.setName(nextSheetName);
    ss.setActiveSheet(newSheet);
    ss.moveActiveSheet(3); // 👈 新シートを左から3番目の位置にピタッと差し込み
    
    ss.toast('3/4: 新シートのデータをクリア中...', '処理中', 10000);
    
    // 6. 新シートの2行目以降のデータをクリア（1行目の見出しや列幅、条件付き書式は維持）
    const lastRow = newSheet.getLastRow();
    const lastColumn = newSheet.getLastColumn();
    
    if (lastRow > 1) {
      newSheet.getRange(2, 1, lastRow - 1, lastColumn).clearContent();
    }
    
    ss.toast('4/4: 見やすさ向上の色付けデザインを適用中...', '処理中', 10000);
    
    // 7. 灰色以外のパステルミントグリーンで行ごと色分け（オルタネートカラー）を適用
    applyAlternatingRowColors(newSheet);
    
    // 完了通知
    ss.toast('新年度への更新がすべて完了しました！', '✨ 成功', 5000);
    SpreadsheetApp.getUi().alert(
      '🎉 更新完了（トリガー自動追従）',
      '年度更新処理が正常に完了しました！\n\n' +
      '① 📦_バックアップ保存箱 に元のデータの複製を保存しました。\n' +
      '② 3番目のタブに新シート「' + nextSheetName + '」を挿入し、2行目を初期化しました。\n' +
      '③ システム全体のロックを解除しました（自動実行トリガーが再開されます）。\n\n' +
      '💡 本体の自動実行トリガーは、設定を変更することなく、そのまま自動的に「' + nextSheetName + '」へ書き込みを始めます。',
      SpreadsheetApp.getUi().ButtonSet.OK
    );

  } catch (error) {
    // エラーが発生した場合も、ダイアログを表示
    SpreadsheetApp.getUi().alert('❌ エラーが発生しました', 'エラー内容: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  } finally {
    // 🔓 8. 処理が成功しても失敗しても、必ずロックを解除（自動実行をONに戻す）
    lock.releaseLock();
  }
}

/**
 * 本体ファイルと同じ階層にバックアップを作成する関数
 */
function createSystemBackup(ss, version) {
  const fileId = ss.getId();
  const file = DriveApp.getFileById(fileId);
  const parentFolders = file.getParents();
  
  if (!parentFolders.hasNext()) {
    throw new Error('ファイルの親フォルダが見つからないため、バックアップを作成できませんでした。');
  }
  
  const parentFolder = parentFolders.next();
  let backupFolder;
  const folderName = SystemConfig.BACKUP_FOLDER_NAME;
  
  const folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    backupFolder = folders.next();
  } else {
    backupFolder = parentFolder.createFolder(folderName);
  }
  
  const formattedDate = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmm');
  const backupFileName = version + '_' + formattedDate + '_' + ss.getName() + '_退避';
  
  file.makeCopy(backupFileName, backupFolder);
}

/**
 * 現在のシート名から次の年度の連番シート名を自動生成する関数
 */
function generateNextSheetName(currentName) {
  const numberMatch = currentName.match(/\d+/);
  if (numberMatch) {
    const currentNum = parseInt(numberMatch[0], 10);
    const nextNum = currentNum + 1;
    return currentName.replace(numberMatch[0], nextNum);
  } else {
    return currentName + '_新年度';
  }
}

/**
 * 灰色以外のパステルカラーで行ごとに色分けを設定する関数
 */
function applyAlternatingRowColors(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  
  if (lastRow <= 1) return;
  
  sheet.getRange(2, 1, lastRow - 1, lastColumn).setBackground('#FFFFFF');
  const maxApplyRow = Math.max(lastRow, 500); 
  
  for (let r = 2; r <= maxApplyRow; r++) {
    const color = (r % 2 === 0) ? SystemConfig.COLORS.ODD : SystemConfig.COLORS.EVEN;
    sheet.getRange(r, 1, 1, lastColumn).setBackground(color);
  }
}