/**
 * 三者面談 予約表 — クラス別PDF出力処理
 * 
 * 各クラスシートの左側（生徒別）を1ページ目、右側（時間枠別）を2ページ目にした
 * 2ページ構成のPDFを作成日時付きでGoogleドライブの「📄_三者面談PDF」フォルダに保存する。
 */

var PDF_FOLDER_NAME = '📄_三者面談PDF';

/**
 * 全クラスのPDFを一括出力する
 */
function exportAllClassesPdf() {
  var classes = getClasses();
  var folder = getOrCreatePdfFolder_();
  var nowStr = Utilities.formatDate(new Date(), TZ, 'yyyy/MM/dd HH:mm');
  var fileDateStr = Utilities.formatDate(new Date(), TZ, 'yyyyMMdd_HHmm');
  var createdFiles = [];

  for (var i = 0; i < classes.length; i++) {
    var clsName = classes[i].name;
    var file = exportSingleClassPdf_(clsName, folder, nowStr, fileDateStr);
    if (file) {
      createdFiles.push(file);
    }
  }

  return {
    count: createdFiles.length,
    folderUrl: folder.getUrl(),
    folderName: PDF_FOLDER_NAME
  };
}

/**
 * 指定したクラス（または現在アクティブなクラス）のPDFを出力する
 */
function exportCurrentClassPdf(targetClsName) {
  var ss = ss_();
  var clsName = targetClsName;
  if (!clsName) {
    var activeSheet = SpreadsheetApp.getActiveSheet();
    var sheetName = activeSheet.getName();
    if (sheetName.startsWith(CLASS_SHEET_PREFIX)) {
      clsName = sheetName.replace(CLASS_SHEET_PREFIX, '');
    } else {
      var classes = getClasses();
      if (classes.length) clsName = classes[0].name;
    }
  }

  var folder = getOrCreatePdfFolder_();
  var nowStr = Utilities.formatDate(new Date(), TZ, 'yyyy/MM/dd HH:mm');
  var fileDateStr = Utilities.formatDate(new Date(), TZ, 'yyyyMMdd_HHmm');

  var file = exportSingleClassPdf_(clsName, folder, nowStr, fileDateStr);
  return {
    className: clsName,
    fileUrl: file.getUrl(),
    fileName: file.getName(),
    folderUrl: folder.getUrl()
  };
}

/**
 * 1クラス分の2ページPDFを作成する内部関数
 */
function exportSingleClassPdf_(clsName, folder, nowStr, fileDateStr) {
  var ss = ss_();
  var sheetName = CLASS_SHEET_PREFIX + clsName;
  var sh = ss.getSheetByName(sheetName);
  if (!sh) return null;

  var lastRow = sh.getLastRow();
  if (lastRow < 2) return null;

  // 一時的なスプレッドシートを作成して2ページ構成を作る
  var tempSs = SpreadsheetApp.create('Temp_PDF_' + clsName + '_' + fileDateStr);
  var tempSsId = tempSs.getId();

  try {
    // 1ページ目: 生徒別 予約状況
    var p1Sheet = tempSs.getSheets()[0];
    p1Sheet.setName('生徒別予約状況');

    // 1行目: 表題と作成日時
    p1Sheet.getRange(1, 1, 1, 6).merge()
      .setValue('【' + clsName + '】三者面談 生徒別予約状況一覧　（作成日時: ' + nowStr + '）')
      .setFontWeight('bold')
      .setFontSize(13)
      .setBackground('#cfe2f3')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    p1Sheet.setRowHeight(1, 35);

    // 左の表（A1:F{lastRow}）のデータとスタイルをコピー
    var leftRange = sh.getRange(1, 1, lastRow, 6);
    var leftValues = leftRange.getValues();
    var leftBgs = leftRange.getBackgrounds();
    var leftWeights = leftRange.getFontWeights();
    var leftAligns = leftRange.getHorizontalAlignments();

    p1Sheet.getRange(2, 1, lastRow, 6).setValues(leftValues);
    p1Sheet.getRange(2, 1, lastRow, 6).setBackgrounds(leftBgs);
    p1Sheet.getRange(2, 1, lastRow, 6).setFontWeights(leftWeights);
    p1Sheet.getRange(2, 1, lastRow, 6).setHorizontalAlignments(leftAligns);
    p1Sheet.getRange(2, 1, lastRow, 6).setBorder(true, true, true, true, true, true, '#b7b7b7', SpreadsheetApp.BorderStyle.SOLID);
    p1Sheet.getRange(2, 1, 1, 6).setBorder(true, true, true, true, true, true, '#5f6368', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    // 列幅設定
    p1Sheet.setColumnWidth(1, 65);  // 出席番号
    p1Sheet.setColumnWidth(2, 120); // 生徒氏名
    p1Sheet.setColumnWidth(3, 80);  // 予約状況
    p1Sheet.setColumnWidth(4, 180); // 予約日時
    p1Sheet.setColumnWidth(5, 120); // 保護者氏名
    p1Sheet.setColumnWidth(6, 220); // 連絡事項

    // 2ページ目: 時間枠別 予約表
    var p2Sheet = tempSs.insertSheet('時間枠別予約表');

    // 1行目: 表題と作成日時
    p2Sheet.getRange(1, 1, 1, 7).merge()
      .setValue('【' + clsName + '】三者面談 時間枠別予約一覧　（作成日時: ' + nowStr + '）')
      .setFontWeight('bold')
      .setFontSize(13)
      .setBackground('#d9d2e9')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    p2Sheet.setRowHeight(1, 35);

    // 右の表（I1:O{lastRow}）のデータとスタイルをコピー
    var rightRange = sh.getRange(1, 9, lastRow, 7);
    var rightValues = rightRange.getValues();
    var rightBgs = rightRange.getBackgrounds();
    var rightWeights = rightRange.getFontWeights();
    var rightAligns = rightRange.getHorizontalAlignments();

    p2Sheet.getRange(2, 1, lastRow, 7).setValues(rightValues);
    p2Sheet.getRange(2, 1, lastRow, 7).setBackgrounds(rightBgs);
    p2Sheet.getRange(2, 1, lastRow, 7).setFontWeights(rightWeights);
    p2Sheet.getRange(2, 1, lastRow, 7).setHorizontalAlignments(rightAligns);
    p2Sheet.getRange(2, 1, lastRow, 7).setBorder(true, true, true, true, true, true, '#b7b7b7', SpreadsheetApp.BorderStyle.SOLID);
    p2Sheet.getRange(2, 1, 1, 7).setBorder(true, true, true, true, true, true, '#5f6368', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    // 日付境界線の設定
    var prevDate = '';
    for (var r = 1; r < rightValues.length; r++) {
      var curDate = String(rightValues[r][0] || '');
      if (curDate && curDate !== prevDate) {
        p2Sheet.getRange(r + 2, 1, 1, 7).setBorder(true, null, null, null, null, null, '#5f6368', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
        prevDate = curDate;
      }
    }

    // 列幅設定
    p2Sheet.setColumnWidth(1, 110); // 日付
    p2Sheet.setColumnWidth(2, 110); // 時間
    p2Sheet.setColumnWidth(3, 70);  // 状態
    p2Sheet.setColumnWidth(4, 70);  // 出席番号
    p2Sheet.setColumnWidth(5, 120); // 生徒氏名
    p2Sheet.setColumnWidth(6, 120); // 保護者氏名
    p2Sheet.setColumnWidth(7, 90);  // 予約コード

    SpreadsheetApp.flush();

    // PDFエクスポートパラメータ（A4横、幅にフィット、全シート出力）
    var url = 'https://docs.google.com/spreadsheets/d/' + tempSsId + '/export?' +
      'exportFormat=pdf&format=pdf' +
      '&size=A4' +
      '&portrait=false' +      // 横向き
      '&fitw=true' +           // 幅に合わせる
      '&gridlines=false' +     // デフォルト枠線はOFF（明示的な罫線を使用）
      '&printtitle=false' +
      '&sheetnames=false' +
      '&fzr=false';

    var token = ScriptApp.getOAuthToken();
    var response = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    });

    var blob = response.getBlob().setName('【' + clsName + '】三者面談予約表_' + fileDateStr + '.pdf');
    var pdfFile = folder.createFile(blob);

    return pdfFile;
  } finally {
    // 一時スプレッドシートを削除
    try {
      DriveApp.getFileById(tempSsId).setTrashed(true);
    } catch (e) {
      console.warn('一時ファイル削除スキップ:', e);
    }
  }
}

/**
 * PDF保存用フォルダを取得または作成
 */
function getOrCreatePdfFolder_() {
  var ssFile = DriveApp.getFileById(SPREADSHEET_ID);
  var parents = ssFile.getParents();
  var parentFolder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();

  var folders = parentFolder.getFoldersByName(PDF_FOLDER_NAME);
  if (folders.hasNext()) {
    return folders.next();
  }
  return parentFolder.createFolder(PDF_FOLDER_NAME);
}
