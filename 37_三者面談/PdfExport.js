/**
 * 三者面談 予約表 — クラス別PDF出力処理
 * 
 * 1. クラス別 予約一覧表（1P: 生徒別、2P: 時間別）
 * 2. 当日用 面談メモ・カルテ付き 進行シート（日別・手書きメモ欄付き A4縦）
 */

var PDF_FOLDER_NAME = '📄_三者面談PDF';

/**
 * 全クラスの予約一覧PDFを一括出力する
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
 * 指定したクラス（または現在アクティブなクラス）の予約一覧PDFを出力する
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
 * 1クラス分の2ページ予約表PDFを作成する内部関数（A4縦レイアウト）
 */
function exportSingleClassPdf_(clsName, folder, nowStr, fileDateStr) {
  var ss = ss_();
  var sheetName = CLASS_SHEET_PREFIX + clsName;
  var sh = ss.getSheetByName(sheetName);
  if (!sh) return null;

  var lastRow = sh.getLastRow();
  if (lastRow < 2) return null;

  var tempSs = SpreadsheetApp.create('Temp_PDF_' + clsName + '_' + fileDateStr);
  var tempSsId = tempSs.getId();

  try {
    // 1ページ目: 生徒別 予約状況
    var p1Sheet = tempSs.getSheets()[0];
    p1Sheet.setName('生徒別予約状況');

    p1Sheet.getRange(1, 1, 1, 6).merge()
      .setValue('【' + clsName + '】三者面談 生徒別予約状況一覧　（作成日時: ' + nowStr + '）')
      .setFontWeight('bold')
      .setFontSize(12)
      .setBackground('#cfe2f3')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    p1Sheet.setRowHeight(1, 32);

    var leftRange = sh.getRange(1, 1, lastRow, 6);
    p1Sheet.getRange(2, 1, lastRow, 6).setValues(leftRange.getValues());
    p1Sheet.getRange(2, 1, lastRow, 6).setBackgrounds(leftRange.getBackgrounds());
    p1Sheet.getRange(2, 1, lastRow, 6).setFontWeights(leftRange.getFontWeights());
    p1Sheet.getRange(2, 1, lastRow, 6).setHorizontalAlignments(leftRange.getHorizontalAlignments());
    p1Sheet.getRange(2, 1, lastRow, 6).setWrap(true).setVerticalAlignment('middle');
    p1Sheet.getRange(2, 1, lastRow, 6).setBorder(true, true, true, true, true, true, '#b7b7b7', SpreadsheetApp.BorderStyle.SOLID);
    p1Sheet.getRange(2, 1, 1, 6).setBorder(true, true, true, true, true, true, '#5f6368', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    p1Sheet.setColumnWidth(1, 45);  // 出席番号
    p1Sheet.setColumnWidth(2, 95);  // 生徒氏名
    p1Sheet.setColumnWidth(3, 60);  // 予約状況
    p1Sheet.setColumnWidth(4, 130); // 予約日時
    p1Sheet.setColumnWidth(5, 95);  // 保護者氏名
    p1Sheet.setColumnWidth(6, 150); // 連絡事項

    // 2ページ目: 時間枠別 予約表
    var p2Sheet = tempSs.insertSheet('時間枠別予約表');

    p2Sheet.getRange(1, 1, 1, 7).merge()
      .setValue('【' + clsName + '】三者面談 時間枠別予約一覧　（作成日時: ' + nowStr + '）')
      .setFontWeight('bold')
      .setFontSize(12)
      .setBackground('#d9d2e9')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    p2Sheet.setRowHeight(1, 32);

    var rightRange = sh.getRange(1, 9, lastRow, 7);
    var rightValues = rightRange.getValues();
    p2Sheet.getRange(2, 1, lastRow, 7).setValues(rightValues);
    p2Sheet.getRange(2, 1, lastRow, 7).setBackgrounds(rightRange.getBackgrounds());
    p2Sheet.getRange(2, 1, lastRow, 7).setFontWeights(rightRange.getFontWeights());
    p2Sheet.getRange(2, 1, lastRow, 7).setHorizontalAlignments(rightRange.getHorizontalAlignments());
    p2Sheet.getRange(2, 1, lastRow, 7).setWrap(true).setVerticalAlignment('middle');
    p2Sheet.getRange(2, 1, lastRow, 7).setBorder(true, true, true, true, true, true, '#b7b7b7', SpreadsheetApp.BorderStyle.SOLID);
    p2Sheet.getRange(2, 1, 1, 7).setBorder(true, true, true, true, true, true, '#5f6368', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    var prevDate = '';
    for (var r = 1; r < rightValues.length; r++) {
      var curDate = String(rightValues[r][0] || '');
      if (curDate && curDate !== prevDate) {
        p2Sheet.getRange(r + 2, 1, 1, 7).setBorder(true, null, null, null, null, null, '#5f6368', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
        prevDate = curDate;
      }
    }

    p2Sheet.setColumnWidth(1, 85); // 日付
    p2Sheet.setColumnWidth(2, 85); // 時間
    p2Sheet.setColumnWidth(3, 50); // 状態
    p2Sheet.setColumnWidth(4, 45); // 出席番号
    p2Sheet.setColumnWidth(5, 95); // 生徒氏名
    p2Sheet.setColumnWidth(6, 95); // 保護者氏名
    p2Sheet.setColumnWidth(7, 65); // 予約コード

    SpreadsheetApp.flush();

    var url = 'https://docs.google.com/spreadsheets/d/' + tempSsId + '/export?' +
      'exportFormat=pdf&format=pdf' +
      '&size=A4' +
      '&portrait=true' +
      '&fitw=true' +
      '&gridlines=false' +
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
    try {
      DriveApp.getFileById(tempSsId).setTrashed(true);
    } catch (e) {
      console.warn('一時ファイル削除スキップ:', e);
    }
  }
}

/**
 * 当日面談メモ付き 進行シート（PDF）を全クラス一括作成
 */
function exportAllMeetingNotesPdf() {
  var classes = getClasses();
  var folder = getOrCreatePdfFolder_();
  var nowStr = Utilities.formatDate(new Date(), TZ, 'yyyy/MM/dd HH:mm');
  var fileDateStr = Utilities.formatDate(new Date(), TZ, 'yyyyMMdd_HHmm');
  var createdFiles = [];

  for (var i = 0; i < classes.length; i++) {
    var clsName = classes[i].name;
    var file = exportSingleMeetingNotesPdf_(clsName, folder, nowStr, fileDateStr);
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
 * 現在のクラスの当日面談メモ付き 進行シート（PDF）を作成
 */
function exportCurrentMeetingNotesPdf(targetClsName) {
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

  var file = exportSingleMeetingNotesPdf_(clsName, folder, nowStr, fileDateStr);
  return {
    className: clsName,
    fileUrl: file.getUrl(),
    fileName: file.getName(),
    folderUrl: folder.getUrl()
  };
}

/**
 * 1クラス分の当日面談メモシートPDFを作成（各実施日ごとに1ページ・ゆったり手書きメモ欄）
 */
function exportSingleMeetingNotesPdf_(clsName, folder, nowStr, fileDateStr) {
  var ss = ss_();
  var days = getDays();
  var slots = readSlots_();
  var classes = getClasses();
  var teacherName = '';
  for (var k = 0; k < classes.length; k++) {
    if (classes[k].name === clsName) teacherName = classes[k].teacher;
  }

  var tempSs = SpreadsheetApp.create('Temp_MeetingNotes_' + clsName + '_' + fileDateStr);
  var tempSsId = tempSs.getId();

  try {
    for (var d = 0; d < days.length; d++) {
      var dayDate = days[d];
      var dLabel = dateLabel_(dayDate);
      var daySlots = [];
      for (var s = 0; s < slots.length; s++) {
        var v = slots[s].v;
        if (String(v[COL.CLASS - 1]) === clsName && ymd_(v[COL.DATE - 1]) === ymd_(dayDate)) {
          daySlots.push(v);
        }
      }

      var sheet = (d === 0) ? tempSs.getSheets()[0] : tempSs.insertSheet('Day_' + (d + 1));
      sheet.setName(dLabel.replace(/[()]/g, '_'));

      // タイトルヘッダー
      sheet.getRange(1, 1, 1, 4).merge()
        .setValue('【' + clsName + '】三者面談 進行記録シート　' + dLabel + (teacherName ? '（担任: ' + teacherName + '）' : ''))
        .setFontWeight('bold')
        .setFontSize(12)
        .setBackground('#d9ead3')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');
      sheet.setRowHeight(1, 30);

      var curRow = 2;
      for (var idx = 0; idx < daySlots.length; idx++) {
        var slot = daySlots[idx];
        var timeStr = slot[COL.START - 1] + '〜' + slot[COL.END - 1];
        var st = String(slot[COL.STATUS - 1]);
        var isBooked = (st === STATUS.BOOKED);
        var stNo = slot[COL.NUMBER - 1] ? slot[COL.NUMBER - 1] + '番' : '';
        var stName = slot[COL.STUDENT - 1] || '';
        var guardian = slot[COL.GUARDIAN - 1] ? slot[COL.GUARDIAN - 1] + ' 様' : '';
        var note = slot[COL.NOTE - 1] || 'なし';

        // 枠ヘッダー行（高さ26px）
        sheet.getRange(curRow, 1, 1, 4).merge()
          .setValue('【第' + (idx + 1) + '枠】 ' + timeStr + '　' + (isBooked ? stNo + ' ' + stName + '（保護者: ' + guardian + '）' : '（※' + st + '）'))
          .setFontWeight('bold')
          .setFontSize(10)
          .setBackground(isBooked ? '#e8f0fe' : '#f1f3f4')
          .setVerticalAlignment('middle');
        sheet.setRowHeight(curRow, 26);
        curRow++;

        // 内容ブロック（左: 予約詳細・連絡事項、右: 手書きメモ欄）
        sheet.getRange(curRow, 1, 3, 2).merge()
          .setValue('■ 事前の連絡・相談事項:\n' + (isBooked ? note : '—'))
          .setFontSize(9)
          .setWrap(true)
          .setVerticalAlignment('top');

        // ［進路・学習面］行（ゆったり42px）
        sheet.getRange(curRow, 3, 1, 2).merge()
          .setValue('［進路・学習面］')
          .setFontSize(9).setFontColor('#5f6368').setVerticalAlignment('top');
        sheet.setRowHeight(curRow, 42);
        curRow++;

        // ［生活・友人・家庭］行（ゆったり42px）
        sheet.getRange(curRow, 3, 1, 2).merge()
          .setValue('［生活・友人・家庭］')
          .setFontSize(9).setFontColor('#5f6368').setVerticalAlignment('top');
        sheet.setRowHeight(curRow, 42);
        curRow++;

        // ［次への確認事項・申し送り］行（ゆったり40px）
        sheet.getRange(curRow, 3, 1, 2).merge()
          .setValue('［次への確認事項・申し送り］')
          .setFontSize(9).setFontColor('#5f6368').setVerticalAlignment('top');
        sheet.setRowHeight(curRow, 40);
        curRow++;

        // 枠を囲む罫線
        sheet.getRange(curRow - 4, 1, 4, 4)
          .setBorder(true, true, true, true, true, true, '#5f6368', SpreadsheetApp.BorderStyle.SOLID);
        
        sheet.setRowHeight(curRow, 10); // 余白行（10px）
        curRow++;
      }

      sheet.setColumnWidth(1, 100);
      sheet.setColumnWidth(2, 130);
      sheet.setColumnWidth(3, 160);
      sheet.setColumnWidth(4, 160);
    }

    SpreadsheetApp.flush();

    var url = 'https://docs.google.com/spreadsheets/d/' + tempSsId + '/export?' +
      'exportFormat=pdf&format=pdf' +
      '&size=A4' +
      '&portrait=true' +
      '&fitw=true' +
      '&gridlines=false' +
      '&printtitle=false' +
      '&sheetnames=false' +
      '&fzr=false';

    var token = ScriptApp.getOAuthToken();
    var response = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    });

    var blob = response.getBlob().setName('【' + clsName + '】三者面談_当日進行記録シート_' + fileDateStr + '.pdf');
    var pdfFile = folder.createFile(blob);

    return pdfFile;
  } finally {
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

/**
 * 単体クラスPDF作成完了ダイアログ（クリック可能なリンクボタン付き）
 */
function showPdfCompleteDialog_(res, isMeetingNotes) {
  var label = isMeetingNotes ? '当日面談記録シート' : '予約表PDF';
  var html = '<!DOCTYPE html><html><head><base target="_blank">' +
    '<style>' +
    'body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; padding: 16px 20px; color: #202124; line-height: 1.5; margin: 0; }' +
    'h2 { font-size: 1.15rem; margin: 0 0 12px; color: #1a73e8; display: flex; align-items: center; gap: 8px; }' +
    'p { margin: 0 0 10px; font-size: 0.95rem; }' +
    '.filename { font-size: 0.88rem; color: #5f6368; background: #f1f3f4; padding: 6px 10px; border-radius: 6px; word-break: break-all; margin: 10px 0 16px; }' +
    '.btn-group { display: flex; flex-direction: column; gap: 10px; margin-top: 18px; }' +
    'a.btn { display: block; text-align: center; padding: 10px 14px; text-decoration: none; border-radius: 6px; font-weight: 600; font-size: 0.95rem; box-sizing: border-box; }' +
    'a.btn-primary { background: #1a73e8; color: #ffffff; }' +
    'a.btn-primary:hover { background: #1557b0; }' +
    'a.btn-secondary { background: #ffffff; color: #1a73e8; border: 1px solid #dadce0; }' +
    'a.btn-secondary:hover { background: #f8f9fa; }' +
    'button.btn-close { width: 100%; padding: 8px; background: transparent; border: none; color: #5f6368; font-size: 0.88rem; cursor: pointer; margin-top: 6px; }' +
    'button.btn-close:hover { text-decoration: underline; }' +
    '</style></head><body>' +
    '<h2>📄 ' + label + 'の作成が完了しました</h2>' +
    '<p><strong>【' + res.className + '】</strong>の' + label + 'を作成しました。</p>' +
    '<div class="filename">📎 ' + res.fileName + '</div>' +
    '<div class="btn-group">' +
    '<a class="btn btn-primary" href="' + res.fileUrl + '">📄 作成したPDFを開く</a>' +
    '<a class="btn btn-secondary" href="' + res.folderUrl + '">📁 保存先フォルダを開く</a>' +
    '<button class="btn-close" onclick="google.script.host.close()">閉じる</button>' +
    '</div></body></html>';

  var userInterface = HtmlService.createHtmlOutput(html)
    .setWidth(420)
    .setHeight(270);
  SpreadsheetApp.getUi().showModalDialog(userInterface, label + '作成完了');
}

/**
 * 全クラスPDF作成完了ダイアログ（クリック可能なリンクボタン付き）
 */
function showAllPdfCompleteDialog_(res, isMeetingNotes) {
  var label = isMeetingNotes ? '当日面談記録シート' : '予約表PDF';
  var html = '<!DOCTYPE html><html><head><base target="_blank">' +
    '<style>' +
    'body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; padding: 16px 20px; color: #202124; line-height: 1.5; margin: 0; }' +
    'h2 { font-size: 1.15rem; margin: 0 0 12px; color: #1a73e8; display: flex; align-items: center; gap: 8px; }' +
    'p { margin: 0 0 10px; font-size: 0.95rem; }' +
    '.folder-box { font-size: 0.9rem; color: #202124; background: #e8f0fe; padding: 8px 12px; border-radius: 6px; margin: 10px 0 16px; }' +
    '.btn-group { display: flex; flex-direction: column; gap: 10px; margin-top: 18px; }' +
    'a.btn { display: block; text-align: center; padding: 10px 14px; text-decoration: none; border-radius: 6px; font-weight: 600; font-size: 0.95rem; box-sizing: border-box; }' +
    'a.btn-primary { background: #1a73e8; color: #ffffff; }' +
    'a.btn-primary:hover { background: #1557b0; }' +
    'button.btn-close { width: 100%; padding: 8px; background: transparent; border: none; color: #5f6368; font-size: 0.88rem; cursor: pointer; margin-top: 6px; }' +
    'button.btn-close:hover { text-decoration: underline; }' +
    '</style></head><body>' +
    '<h2>📄 全クラス ' + label + '作成完了</h2>' +
    '<p>全 <strong>' + res.count + '</strong> クラス分の' + label + 'を作成しました。</p>' +
    '<div class="folder-box">📁 保存先: <strong>' + res.folderName + '</strong> フォルダ</div>' +
    '<div class="btn-group">' +
    '<a class="btn btn-primary" href="' + res.folderUrl + '">📁 保存先フォルダを開く</a>' +
    '<button class="btn-close" onclick="google.script.host.close()">閉じる</button>' +
    '</div></body></html>';

  var userInterface = HtmlService.createHtmlOutput(html)
    .setWidth(420)
    .setHeight(250);
  SpreadsheetApp.getUi().showModalDialog(userInterface, '全クラス' + label + '作成完了');
}
