/**
 * @file Main.gs
 * @description 教育現場用PDF連携・管理システムのメイン処理ロジック。
 * @version v1.1.0
 */

/**
 * すべてのプロセス（①〜⑤）を順番に一括実行します。
 */
function executeAllProcesses() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    '一括実行の確認',
    'ファイル名の生成(K列)から照合まで、すべての処理を開始しますか？\n（※システムバージョン: ' + Config.VERSION + '）',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;

  // 処理開始時にトースト通知で安心感を提供
  SpreadsheetApp.getActiveSpreadsheet().toast('全工程の一括実行を開始します...', '処理中', -1);

  try {
    // 💡 安全対策：実行前に自動バックアップを作成
    createBackupOfCurrentFile();

    generateFileNamesFromEmailAndName(); // 1. ファイル名生成
    renamePdfExtensions();               // 2. 拡張子の統一
    getFileNamesAndUrlsFromFolder();     // 3. フォルダ情報の取得
    setPdfUrls();                        // 4. 一覧シートへのURL書き込み
    checkFileUrlMatches();               // 5. 照合チェック
    
    SpreadsheetApp.getActiveSpreadsheet().toast('すべての工程が正常に完了しました！', '成功', 5);
    Browser.msgBox('✨ すべての工程が完了しました！\n※「' + Config.BACKUP_FOLDER_NAME + '」に実行前のバックアップを保存しました。');
    
  } catch (e) {
    Logger.log(e.stack);
    Browser.msgBox('一括実行中にエラーが発生しました:\n' + e.message);
  }
}

/**
 * ① F列(氏名)とJ列(メール)から、K列(ファイル名)を生成する関数
 */
function generateFileNamesFromEmailAndName() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(Config.SHEET_NAMES.LIST);

  if (!sheet) {
    Browser.msgBox('エラー: 「' + Config.SHEET_NAMES.LIST + '」シートが見つかりません。');
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Browser.msgBox('データがありません。');
    return;
  }

  // 氏名(F列)からメールアドレス(J列)までの範囲を一括取得
  const startCol = Config.LIST_COL_NAME;
  const numCols = Config.LIST_COL_EMAIL - startCol + 1;
  const data = sheet.getRange(2, startCol, lastRow - 1, numCols).getValues();
  
  const newFileNames = [];
  let count = 0;

  for (let i = 0; i < data.length; i++) {
    const nameRaw = data[i][0]; // F列にあたる位置
    const emailRaw = data[i][data[i].length - 1]; // J列にあたる位置
    
    if (nameRaw && emailRaw) {
      // 💡 揺らぎ補正：氏名に含まれる全角・半角スペースを除去して結合
      const cleanName = String(nameRaw).replace(/[\s ]/g, '');
      // メールアドレスの@より前（ユーザID部分）を取得
      const emailPrefix = String(emailRaw).split('@')[0].trim();
      
      // 生成ルール: [メール前部]_[氏名].PDF
      const fileName = emailPrefix + '_' + cleanName + '.PDF';
      newFileNames.push([fileName]);
      count++;
    } else {
      newFileNames.push(['']);
    }
  }

  // K列(ファイル名)に一括書き込み
  sheet.getRange(2, Config.LIST_COL_FILENAME, newFileNames.length, 1).setValues(newFileNames);

  // 一括実行から呼ばれたのではない場合のみ完了メッセージを表示
  const caller = new Error().stack;
  if (!caller || !caller.includes('executeAllProcesses')) {
    Browser.msgBox('完了: K列にファイル名を生成しました。(' + count + '件)\n※名前のスペースは自動で整えられました。');
  }
}

/**
 * ② Googleドライブ内PDFの拡張子を大文字(.PDF)に統一する関数
 */
function renamePdfExtensions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(Config.SHEET_NAMES.BASE);
  if (!sheet) return;

  const folderUrlOrId = sheet.getRange(Config.BASE_CELL_FOLDER_URL).getValue();
  const folderId = extractFolderId(folderUrlOrId);
  if (!folderId) {
    Browser.msgBox('エラー: 基礎データシートのB2セルからフォルダIDを判別できませんでした。');
    return;
  }

  try {
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFiles();
    let count = 0;

    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();
      // 小文字の「.pdf」で終わるファイルがあれば大文字に置換
      if (fileName.endsWith('.pdf')) {
        const newName = fileName.slice(0, -4) + '.PDF';
        file.setName(newName);
        count++;
      }
    }
    
    const caller = new Error().stack;
    if (!caller || !caller.includes('executeAllProcesses')) {
      Browser.msgBox("拡張子統一完了: " + count + "件のファイルを小文字から大文字(.PDF)に変換しました。");
    }
  } catch (e) {
    Browser.msgBox("拡張子変更中にエラーが発生しました: " + e.message);
  }
}

/**
 * ③ 指定フォルダ内の全ファイル名とURLを「フォルダ」シートに抽出する関数
 */
function getFileNamesAndUrlsFromFolder() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const baseDataSheet = spreadsheet.getSheetByName(Config.SHEET_NAMES.BASE);
  const folderSheet = spreadsheet.getSheetByName(Config.SHEET_NAMES.FOLDER);
  
  if (!baseDataSheet || !folderSheet) return;
  
  const lastRow = folderSheet.getLastRow();
  // 古い既存データをクリア（A2:Dまで）
  if (lastRow >= 2) folderSheet.getRange(2, 1, lastRow - 1, 4).clearContent();
  
  const folderUrlOrId = baseDataSheet.getRange(Config.BASE_CELL_FOLDER_URL).getValue();
  const folderId = extractFolderId(folderUrlOrId);
  if (!folderId) return;
  
  let fileData = [];
  
  try {
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFiles();
    let count = 1;
    
    while (files.hasNext()) {
      const file = files.next();
      fileData.push([count, file.getName(), file.getUrl(), '']);
      count++;
    }
  } catch(e) {
    Browser.msgBox('エラー: Googleドライブのフォルダにアクセスできませんでした。');
    return;
  }
  
  if (fileData.length > 0) {
    folderSheet.getRange(2, 1, fileData.length, 4).setValues(fileData);
    
    const caller = new Error().stack;
    if (!caller || !caller.includes('executeAllProcesses')) {
      Browser.msgBox('フォルダリスト更新完了: ' + fileData.length + ' 件のファイルを「フォルダ」シートに洗い出しました。');
    }
  }
}

/**
 * ④ 「フォルダ」シートの情報に基づき、「一覧」シートのK列(ファイル名)と一致するURLをL列に書き込む関数
 */
function setPdfUrls() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const baseSheet = ss.getSheetByName(Config.SHEET_NAMES.BASE);
  const listSheet = ss.getSheetByName(Config.SHEET_NAMES.LIST);

  if (!baseSheet || !listSheet) {
    Browser.msgBox('エラー: 必要となるシート（基礎データ または 一覧）が見つかりません。');
    return;
  }

  const folderUrlOrId = baseSheet.getRange(Config.BASE_CELL_FOLDER_URL).getValue();
  const folderId = extractFolderId(folderUrlOrId);
  if (!folderId) {
    Browser.msgBox('エラー: フォルダIDの取得に失敗しました。');
    return;
  }

  try {
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFilesByType(MimeType.PDF);
    const fileMap = {};
    let totalPdfCount = 0;

    // ドライブ内のPDF名称とURLの対応マップを作成
    while (files.hasNext()) {
      const file = files.next();
      // 💡 揺らぎ補正：ドライブ内のファイル名からも念のためスペースを抜いてキーにする
      const cleanFileName = file.getName().replace(/[\s ]/g, '');
      fileMap[cleanFileName] = file.getUrl();
      totalPdfCount++;
    }

    const lastRow = listSheet.getLastRow();
    if (lastRow < 2) return;

    // K列（ファイル名）のデータを一括取得
    const fileNames = listSheet.getRange(2, Config.LIST_COL_FILENAME, lastRow - 1, 1).getValues();
    const urls = [];
    let matchCount = 0;

    for (let i = 0; i < fileNames.length; i++) {
      const targetNameRaw = fileNames[i][0];
      const cleanTargetName = String(targetNameRaw).replace(/[\s ]/g, '');
      
      if (cleanTargetName && fileMap[cleanTargetName]) {
        urls.push([fileMap[cleanTargetName]]);
        matchCount++;
      } else {
        urls.push(['']);
      }
    }

    // L列（URL）へ一括書き込み
    listSheet.getRange(2, Config.LIST_COL_URL, urls.length, 1).setValues(urls);

    const caller = new Error().stack;
    if (!caller || !caller.includes('executeAllProcesses')) {
      Browser.msgBox('処理完了！\n\n📂 フォルダ内PDF総数: ' + totalPdfCount + ' 件\n✅ リンク紐付け完了(L列): ' + matchCount + ' 件');
    }

  } catch (e) {
    Browser.msgBox('URLの書き込み中にエラーが発生しました: ' + e.message);
  }
}

/**
 * ⑤ 「一覧」シートに紐付いたURLをベースに、「フォルダ」シート側のファイルが網羅されているかチェックする関数
 */
function checkFileUrlMatches() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const listSheet = spreadsheet.getSheetByName(Config.SHEET_NAMES.LIST);
  const folderSheet = spreadsheet.getSheetByName(Config.SHEET_NAMES.FOLDER);
  
  if (!listSheet || !folderSheet) return;

  const folderLastRow = folderSheet.getLastRow();
  if (folderLastRow >= 2) {
    folderSheet.getRange(2, Config.FOLDER_COL_CHECK, folderLastRow - 1, 1).clearContent();
  }

  const listLastRow = listSheet.getLastRow();
  if (listLastRow < 2) return;
  
  // 一覧シートのL列(URL)を取得
  const listUrls = listSheet.getRange(2, Config.LIST_COL_URL, listLastRow - 1, 1).getValues();
  const referenceUrls = new Set();
  
  listUrls.forEach(row => {
    const url = row[0].toString().trim();
    if (url) referenceUrls.add(url);
  });

  if (folderLastRow < 2) return;
  
  // フォルダシートのC列(URL)を取得して照合
  const folderUrls = folderSheet.getRange(2, Config.FOLDER_COL_URL, folderLastRow - 1, 1).getValues();
  const matchResults = [];
  let countMatch = 0;
  let countNoMatch = 0;

  folderUrls.forEach(row => {
    const url = row[0].toString().trim();
    if (url && referenceUrls.has(url)) {
      matchResults.push(['〇']);
      countMatch++;
    } else {
      matchResults.push(['✖']);
      countNoMatch++;
    }
  });

  // フォルダシートのD列に結果を書き込み
  folderSheet.getRange(2, Config.FOLDER_COL_CHECK, matchResults.length, 1).setValues(matchResults);

  const caller = new Error().stack;
  if (!caller || !caller.includes('executeAllProcesses')) {
    Browser.msgBox('照合完了\n\n✅ 名簿と一致したファイル: ' + countMatch + ' 件\n❌ 名簿にない不明なファイル: ' + countNoMatch + ' 件');
  }
}

/**
 * 🛠️ 共通便利ロジック：URLからフォルダIDを安全に切り出す、またはIDをそのまま返す関数
 */
function extractFolderId(input) {
  if (!input) return null;
  const strInput = String(input).trim();
  
  const match = strInput.match(/folders\/([a-zA-Z0-9_-]+)/);
  if (match && match[1]) {
    return match[1];
  }
  
  if (!strInput.includes('http') && strInput.length > 20) {
    return strInput.split('?')[0];
  }
  
  return null;
}

/**
 * 🛠️ 共通便利ロジック：一括実行の直前に本体と同じ階層へバックアップファイルを自動生成する関数
 */
function createBackupOfCurrentFile() {
  try {
    const currentFileId = SpreadsheetApp.getActiveSpreadsheet().getId();
    const currentFile = DriveApp.getFileById(currentFileId);
    const parentFolders = currentFile.getParents();
    
    if (!parentFolders.hasNext()) return;
    const parentFolder = parentFolders.next();
    
    const subFolders = parentFolder.getFoldersByName(Config.BACKUP_FOLDER_NAME);
    let backupFolder;
    if (subFolders.hasNext()) {
      backupFolder = subFolders.next();
    } else {
      backupFolder = parentFolder.createFolder(Config.BACKUP_FOLDER_NAME);
    }
    
    const formattedDate = Utilities.formatDate(new Date(), "JST", "yyyyMMdd_HHmmss");
    const backupName = Config.VERSION + '_[' + formattedDate + ']_' + currentFile.getName();
    
    currentFile.makeCopy(backupName, backupFolder);
  } catch (e) {
    Logger.log('バックアップ作成に失敗しました: ' + e.message);
  }
}