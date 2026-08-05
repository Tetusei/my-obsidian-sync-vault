/**
 * ==============================================================================
 * Googleスプレッドシート PDF分割＆リネーム＆フォルダ移動＆バックアップ ツール (GAS)
 * ==============================================================================
 * 
 * 【最新のセル・列配置】
 * - A2セル: 分割対象PDFのGoogleドライブURLまたはID
 * - B2セル: 分割ページ数（例: 1）
 * - C2セル: 分割ファイルの一時保存先フォルダURLまたはID
 * - D列(D2:D): 分割されたPDFファイルのURL出力先
 * - E列(E2:E): アカウント名
 * - F列(F2:F): 名前
 * - G列(G2:G): リネーム後のファイル名（数式・関数セル。※上書きせず保持）
 * - H列(H2:H): リネーム後、GASによりHYPERLINK（ファイルリンク）が出力される
 * - I2セル: 最終移動先フォルダのURLまたはID
 */

// GAS環境向けタイマー・Web APIポリフィル
(function(g) {
  if (typeof g.setTimeout === 'undefined') {
    g.setTimeout = function(fn, delay, ...args) {
      if (delay && delay > 0) {
        Utilities.sleep(delay);
      }
      try {
        return fn(...args);
      } catch (e) {
        console.error('setTimeout callback error:', e);
      }
    };
  }
  if (typeof g.clearTimeout === 'undefined') {
    g.clearTimeout = function() {};
  }
  if (typeof g.setInterval === 'undefined') {
    g.setInterval = function(fn, delay, ...args) {
      if (delay && delay > 0) {
        Utilities.sleep(delay);
      }
      return fn(...args);
    };
  }
  if (typeof g.clearInterval === 'undefined') {
    g.clearInterval = function() {};
  }
  if (typeof g.performance === 'undefined') {
    g.performance = { now: function() { return Date.now(); } };
  }
})(typeof globalThis !== 'undefined' ? globalThis : this);

const SPREADSHEET_ID = '17_RJte0dNEwc21Rfe5ii5bbYhwrH6-RNKA6ys-9EOBk';

/**
 * アクティブなスプレッドシートを取得（UI非依存）
 */
function getTargetSpreadsheet() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  return ss;
}

/**
 * スプレッドシートが開いたときにカスタムメニューを追加（親しみやすい絵文字アイコン付き）
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📄 PDF分割＆整理ツール')
    .addItem('✂️ 1. PDFを分割してD列にURL出力 (C2フォルダ)', 'splitPdfAndWriteUrls')
    .addItem('🏷️ 2. G列のファイル名でリネーム ➔ H列にリンク作成', 'renameSplitFilesAndAddLinks')
    .addItem('🚚 3. I2セルのフォルダへファイルを移動', 'moveFilesToI2Folder')
    .addSeparator()
    .addItem('⚡ 【一括実行】分割 ➔ リネーム＆リンク ➔ フォルダ移動', 'executeAllSteps')
    .addSeparator()
    .addItem('💾 スプレッドシートをバックアップ', 'backupSpreadsheet')
    .addToUi();
}

/**
 * pdf-lib ライブラリを動的に読み込む
 */
function getPdfLib() {
  if (typeof PDFLib !== 'undefined') {
    return PDFLib;
  }
  if (typeof globalThis.PDFLib !== 'undefined') {
    return globalThis.PDFLib;
  }
  const cdnUrl = 'https://cdnjs.cloudflare.com/ajax/libs/pdf-lib/1.17.1/pdf-lib.min.js';
  const response = UrlFetchApp.fetch(cdnUrl);
  const scriptText = response.getContentText();
  (0, eval)(scriptText);
  return typeof PDFLib !== 'undefined' ? PDFLib : globalThis.PDFLib;
}

/**
 * URLまたは文字列からGoogleドライブのファイルIDを抽出する
 */
function extractFileId(input) {
  if (!input) return null;
  const str = input.toString().trim();
  const match = str.match(/[-\w]{25,}/);
  return match ? match[0] : str;
}

/**
 * URLまたは文字列からGoogleドライブのフォルダIDを抽出する
 */
function extractFolderId(input) {
  if (!input) return null;
  const str = input.toString().trim();
  const folderMatch = str.match(/folders\/([-\w]{25,})/);
  if (folderMatch) {
    return folderMatch[1];
  }
  const idMatch = str.match(/[-\w]{25,}/);
  return idMatch ? idMatch[0] : str;
}

/**
 * スプレッドシートが格納されている親フォルダを取得する
 */
function getSpreadsheetFolder() {
  const ss = getTargetSpreadsheet();
  const file = DriveApp.getFileById(ss.getId());
  const parents = file.getParents();
  if (parents.hasNext()) {
    return parents.next();
  }
  return DriveApp.getRootFolder();
}

/**
 * 1. A2セルのPDFをB2セルのページ数ごとに分割し、C2セルのフォルダに保存してD列にURLを出力する
 */
async function splitPdfAndWriteUrls() {
  const ss = getTargetSpreadsheet();
  const sheet = ss.getActiveSheet();
  const pdfUrlOrId = sheet.getRange('A2').getValue();
  const pagesPerSplit = parseInt(sheet.getRange('B2').getValue(), 10);
  const destFolderUrlOrId = sheet.getRange('C2').getValue();

  if (!pdfUrlOrId) {
    showAlertOrLog('⚠️ エラー: A2セルにPDFのファイルURLまたはIDが指定されていません。');
    return;
  }

  if (isNaN(pagesPerSplit) || pagesPerSplit <= 0) {
    showAlertOrLog('⚠️ エラー: B2セルに有効な分割ページ数（1以上の数値）を指定してください。');
    return;
  }

  // C2セルのフォルダを取得（指定がない場合はスプレッドシートと同じフォルダ）
  let destFolder;
  if (destFolderUrlOrId) {
    const folderId = extractFolderId(destFolderUrlOrId);
    try {
      destFolder = DriveApp.getFolderById(folderId);
    } catch (e) {
      showAlertOrLog('⚠️ エラー: C2セルの保存先フォルダを取得できませんでした。\n' + e.message);
      return;
    }
  } else {
    destFolder = getSpreadsheetFolder();
  }

  const fileId = extractFileId(pdfUrlOrId);
  let srcFile;
  try {
    srcFile = DriveApp.getFileById(fileId);
  } catch (e) {
    showAlertOrLog('⚠️ エラー: A2セルのPDFファイルを取得できませんでした。\n' + e.message);
    return;
  }

  showToastOrLog('pdf-libライブラリを読み込んでいます...', '⏳ 処理中');
  const PDFLib = getPdfLib();

  showToastOrLog('PDFを解析して分割処理を行っています...', '⏳ 処理中');
  
  try {
    const rawBytes = srcFile.getBlob().getBytes();
    const uint8Array = new Uint8Array(rawBytes.length);
    for (let i = 0; i < rawBytes.length; i++) {
      uint8Array[i] = rawBytes[i] < 0 ? rawBytes[i] + 256 : rawBytes[i];
    }

    const srcDoc = await PDFLib.PDFDocument.load(uint8Array);
    const totalPages = srcDoc.getPageCount();
    const baseName = srcFile.getName().replace(/\.pdf$/i, '');

    const newUrls = [];

    for (let startPage = 0; startPage < totalPages; startPage += pagesPerSplit) {
      const endPage = Math.min(startPage + pagesPerSplit, totalPages);
      const subDoc = await PDFLib.PDFDocument.create();
      
      const pageIndices = [];
      for (let i = startPage; i < endPage; i++) {
        pageIndices.push(i);
      }

      const copiedPages = await subDoc.copyPages(srcDoc, pageIndices);
      copiedPages.forEach(p => subDoc.addPage(p));

      const subPdfBytes = await subDoc.save();
      const partIndex = Math.floor(startPage / pagesPerSplit) + 1;
      const splitFileName = `${baseName}_part_${partIndex}.pdf`;
      const blob = Utilities.newBlob(Array.from(subPdfBytes), 'application/pdf', splitFileName);
      
      const splitFile = destFolder.createFile(blob);
      newUrls.push([splitFile.getUrl()]);
    }

    // D2セル以降をクリア
    const lastRow = Math.max(sheet.getLastRow(), 2);
    sheet.getRange(2, 4, lastRow - 1, 1).clearContent(); // D2:D列クリア

    // 分割後のURLを書き出し
    if (newUrls.length > 0) {
      sheet.getRange(2, 4, newUrls.length, 1).setValues(newUrls);
    }

    showToastOrLog(`PDFの分割完了: ${newUrls.length}件のファイルを作成しました。`, '✅ 完了');
    showAlertOrLog(`🎉【分割完了】\n\n・保存先フォルダ: ${destFolder.getName()}\n・出力件数: ${newUrls.length}件（D列に出力）`);

  } catch (err) {
    showAlertOrLog('⚠️ PDF分割処理中にエラーが発生しました:\n' + err.message);
  }
}

/**
 * 2. D列のファイルURLを参照し、G列のファイル名に変更＆H列セルにハイパーリンクを書き出す（G列の数式は保持）
 */
function renameSplitFilesAndAddLinks() {
  const ss = getTargetSpreadsheet();
  const sheet = ss.getActiveSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    showAlertOrLog('⚠️ エラー: 処理対象のデータ（D列・G列）が存在しません。');
    return;
  }

  const dValues = sheet.getRange(2, 4, lastRow - 1, 1).getValues(); // D2:D
  const gValues = sheet.getRange(2, 7, lastRow - 1, 1).getValues(); // G2:G (数式の計算結果)

  // H列クリア
  sheet.getRange(2, 8, lastRow - 1, 1).clearContent();

  let successCount = 0;
  let errorCount = 0;

  for (let i = 0; i < dValues.length; i++) {
    const fileUrl = dValues[i][0];
    let newName = gValues[i][0];

    if (!fileUrl || !newName) continue;

    newName = newName.toString().trim();
    if (!newName) continue;

    if (!/\.pdf$/i.test(newName)) {
      newName += '.pdf';
    }

    const fileId = extractFileId(fileUrl);
    if (!fileId) continue;

    try {
      // 1. ドライブ上のファイル名をリネーム
      const file = DriveApp.getFileById(fileId);
      file.setName(newName);

      // 2. G列の数式は保護し、H列(Column 8)にハイパーリンク付きで出力
      const targetCellH = sheet.getRange(i + 2, 8);
      const escapedName = newName.replace(/"/g, '""');
      targetCellH.setFormula(`=HYPERLINK("${fileUrl}", "${escapedName}")`);

      successCount++;
    } catch (e) {
      console.error(`行 ${i + 2} のリネーム・H列リンク失敗: ` + e.message);
      errorCount++;
    }
  }

  if (successCount === 0 && errorCount === 0) {
    showAlertOrLog('⚠️ 変更対象のデータが見つかりませんでした。');
  } else {
    showAlertOrLog(`🎉【リネーム＆リンク出力完了】\n\n・G列の数式を維持したまま、H列にハイパーリンクを出力しました。\n・成功: ${successCount}件\n・失敗: ${errorCount}件`);
  }
}

/**
 * 3. I2セルに指定されたフォルダURL/IDのフォルダへD列の分割ファイルを移動する
 */
function moveFilesToI2Folder() {
  const ss = getTargetSpreadsheet();
  const sheet = ss.getActiveSheet();
  const folderUrlOrId = sheet.getRange('I2').getValue();

  if (!folderUrlOrId) {
    showAlertOrLog('⚠️ エラー: I2セルに移動先フォルダのURLまたはIDが入力されていません。');
    return;
  }

  const folderId = extractFolderId(folderUrlOrId);
  let targetFolder;

  try {
    targetFolder = DriveApp.getFolderById(folderId);
  } catch (e) {
    showAlertOrLog('⚠️ エラー: I2セルのフォルダを取得できませんでした。\n' + e.message);
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    showAlertOrLog('⚠️ 移動対象のファイル（D列）が存在しません。');
    return;
  }

  const dValues = sheet.getRange(2, 4, lastRow - 1, 1).getValues();
  let moveCount = 0;
  let failCount = 0;

  for (let i = 0; i < dValues.length; i++) {
    const fileUrl = dValues[i][0];
    if (!fileUrl) continue;

    const fileId = extractFileId(fileUrl);
    if (!fileId) continue;

    try {
      const file = DriveApp.getFileById(fileId);
      file.moveTo(targetFolder);
      moveCount++;
    } catch (e) {
      console.error(`行 ${i + 2} のファイル移動失敗: ` + e.message);
      failCount++;
    }
  }

  showAlertOrLog(`🚚【フォルダ移動完了】\n\n・移動先: ${targetFolder.getName()}\n・成功: ${moveCount}件\n・失敗: ${failCount}件`);
}

/**
 * 全ステップを一括実行する（分割 ➔ リネーム＆H列リンク ➔ I2フォルダ移動）
 */
async function executeAllSteps() {
  await splitPdfAndWriteUrls();
  renameSplitFilesAndAddLinks();
  moveFilesToI2Folder();
}

/**
 * スプレッドシートのバックアップを作成する
 */
function backupSpreadsheet() {
  const ss = getTargetSpreadsheet();
  const ssFile = DriveApp.getFileById(ss.getId());
  const parentFolder = getSpreadsheetFolder();

  const folderName = 'バックアップ';
  const folders = parentFolder.getFoldersByName(folderName);
  let backupFolder;

  if (folders.hasNext()) {
    backupFolder = folders.next();
  } else {
    backupFolder = parentFolder.createFolder(folderName);
  }

  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  const backupName = `${ss.getName()}_バックアップ_${timestamp}`;

  ssFile.makeCopy(backupName, backupFolder);

  showAlertOrLog(`💾【バックアップ完了】\n\n・保存先フォルダ: ${backupFolder.getName()}\n・ファイル名: ${backupName}`);
}

/**
 * UIメッセージ出力ヘルパー
 */
function showAlertOrLog(msg) {
  try {
    SpreadsheetApp.getUi().alert(msg);
  } catch (e) {
    console.log('[ALERT] ' + msg);
  }
}

function showToastOrLog(msg, title) {
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(msg, title, 5);
  } catch (e) {
    console.log(`[TOAST] ${title}: ${msg}`);
  }
}
