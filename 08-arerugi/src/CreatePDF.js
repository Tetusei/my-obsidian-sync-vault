/**
 * CreatePDF.gs (v21.0.0：進行状況のトースト通知を追加)
 * 各処理のフェーズごとに、画面右下にトースト（通知）を表示して
 * ユーザーに現在の進行状況を知らせます。
 */

function generateParentForms() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 🌟【トースト】処理開始
  ss.toast("PDFの作成準備を開始します...", "🚀 処理開始", 3);
  
  // 1. 各シートを取得
  const memberSheet = ss.getSheetByName("個人アレルゲン");
  const templateSheet = ss.getSheetByName("保護者記入表"); 
  
  if (!memberSheet) {
    SpreadsheetApp.getUi().alert("「個人アレルゲン」シートが見つかりません。");
    return;
  }
  if (!templateSheet) {
    SpreadsheetApp.getUi().alert("「保護者記入表」シートが見つかりません。");
    return;
  }
  
  // このスプレッドシートと同じフォルダを自動取得
  let parentFolder;
  try {
    const ssId = ss.getId();
    const ssFile = DriveApp.getFileById(ssId);
    parentFolder = ssFile.getParents().next();
  } catch(e) {
    SpreadsheetApp.getUi().alert("このスプレッドシートが保存されているフォルダを特定できませんでした。");
    return;
  }
  
  // 「提出用PDF」フォルダを作成・取得
  let targetFolder;
  const existingFolders = parentFolder.getFoldersByName("提出用PDF");
  if (existingFolders.hasNext()) {
    targetFolder = existingFolders.next();
  } else {
    targetFolder = parentFolder.createFolder("提出用PDF");
  }
  
  // 🌟【トースト】一時ファイルの作成
  ss.toast("一時ファイルを作成し、テンプレートを準備中...", "⏳ 準備中 (1/4)", 4);
  
  // 2. 作業用の一時ファイル（スプレッドシート）を1つ作成
  const formattedDate = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyyMMdd");
  const tempFormName = "【印刷用一括データ】保護者記入表_" + formattedDate;
  
  const oldFiles = targetFolder.getFilesByName(tempFormName);
  while (oldFiles.hasNext()) { oldFiles.next().setTrashed(true); }
  const oldPdfs = targetFolder.getFilesByName(tempFormName + ".pdf");
  while (oldPdfs.hasNext()) { oldPdfs.next().setTrashed(true); }
  
  const tempSpreadsheet = SpreadsheetApp.create(tempFormName);
  const tempFile = DriveApp.getFileById(tempSpreadsheet.getId());
  targetFolder.addFile(tempFile);
  DriveApp.getRootFolder().removeFile(tempFile);
  
  // 一時ファイルにテンプレートをコピーし、「原本」として保護する
  const pristineTemplate = templateSheet.copyTo(tempSpreadsheet);
  pristineTemplate.setName("原本_さわらない");
  
  // 最初から入っていた「シート1（白紙）」をここで確実に削除
  tempSpreadsheet.deleteSheet(tempSpreadsheet.getSheets()[0]);
  SpreadsheetApp.flush(); 
  
  // 🌟【トースト】生徒データの書き込み開始
  ss.toast("生徒ごとのシートを作成し、データを書き込んでいます...", "📝 データ作成中 (2/4)", 5);
  
  // 3. 名簿データの読み込みと置換ループ
  const memberData = memberSheet.getDataRange().getValues();
  
  // F列(インデックス5)・G列(インデックス6)から「毎日対応ルール」を自動学習する
  const dailyMaster = {};
  for (let i = 0; i < memberData.length; i++) {
    const item = memberData[i][5];   // F列
    const action = memberData[i][6]; // G列
    if (item && action) {
      dailyMaster[item.toString().trim()] = action.toString().trim();
    }
  }
  
  let sheetCount = 0;
  
  // データは2行目（インデックス1）から始まる
  for (let i = 1; i < memberData.length; i++) {
    const row = memberData[i];
    const className = row[0];   // A列: クラス
    const studentName = row[1]; // B列: 氏名
    const allergen = row[2];    // C列: アレルゲン
    const dailyItem = row[3];   // D列: 毎日
    
    if (!studentName || studentName === "氏名") continue;
    
    const targetSheet = pristineTemplate.copyTo(tempSpreadsheet);
    targetSheet.setName((sheetCount + 1) + "_" + studentName.substring(0, 10));
    sheetCount++;
    
    targetSheet.createTextFinder("{クラス}").replaceAllWith(className);
    targetSheet.createTextFinder("{氏名}").replaceAllWith(studentName);
    targetSheet.createTextFinder("{アレルゲン}").replaceAllWith(allergen);
    
    // 14行目〜18行目を一旦きれいにリセット
    for (let r = 14; r <= 18; r++) {
      targetSheet.getRange("A" + r).setValue("");
      targetSheet.getRange("C" + r).setValue("");
      targetSheet.getRange("F" + r).setValue("");
    }

    if (dailyItem) {
      const dailyStr = dailyItem.toString();
      const items = dailyStr.split(/[,、・\n]/).map(s => s.trim()).filter(s => s !== "");
      
      let currentRow = 14; // 14行目から書き込みスタート

      for (const item of items) {
        if (currentRow > 18) break; 
        
        const action = dailyMaster[item] || "食べません";
        
        targetSheet.getRange("A" + currentRow).setValue("毎日");
        targetSheet.getRange("C" + currentRow).setValue(item);
        targetSheet.getRange("F" + currentRow).setValue(action);
        currentRow++;
      }
    }
  }
  
  if (sheetCount === 0) {
    SpreadsheetApp.getUi().alert("対象となる生徒のデータがありませんでした。");
    tempFile.setTrashed(true);
    return;
  }
  
  tempSpreadsheet.deleteSheet(pristineTemplate);
  
  // 🌟【トースト】待機と書き込み確定
  ss.toast("文字の書き込みが完了しました。\n安全にPDF化するため、Googleのサーバー処理を5秒間待機します...", "⏳ PDF化準備 (3/4)", 6);
  
  SpreadsheetApp.flush();
  Utilities.sleep(5000); 
  
  // 🌟【トースト】PDF出力開始
  ss.toast("PDFを出力しています！まもなく完了します...", "🖨️ 出力中 (4/4)", -1); // -1を指定すると完了画面が出るまで消えません
  
  // 4. PDFとして書き出し
  const token = ScriptApp.getOAuthToken();
  const pdfUrlExport = "https://docs.google.com/spreadsheets/d/" + tempSpreadsheet.getId() + "/export?" + 
                 "exportFormat=pdf&format=pdf&size=A4&portrait=true&fitw=true&gridlines=false";
  
  const response = UrlFetchApp.fetch(pdfUrlExport, {
    headers: { 'Authorization': 'Bearer ' + token },
    muteHttpExceptions: true
  });
  
  let finalPdfFile;
  if (response.getResponseCode() === 200) {
    const blob = response.getBlob().setName(tempFormName + ".pdf");
    finalPdfFile = targetFolder.createFile(blob); 
  } else {
    SpreadsheetApp.getUi().alert("PDFの書き出し処理に失敗しました。");
    return;
  }
  
  tempFile.setTrashed(true);
  
  const pdfUrl = finalPdfFile.getUrl();
  const folderUrl = targetFolder.getUrl(); 
  
  // 🌟【トースト】完了（不要なトーストを消すために空のトーストを送るか、完了メッセージを出す）
  ss.toast("すべての処理が完了しました！", "✨ 完了", 3);
  
  // 6. 完了画面表示
  const htmlOutput = HtmlService.createHtmlOutput(`
    <div style="font-family: sans-serif; padding: 10px; text-align: center;">
      <h3 style="color: #1a73e8;">✅ PDFの作成が完了しました！</h3>
      <p style="font-size: 14px; color: #555;">ボタンからファイルや保存先を確認できます。</p>
      <br>
      <div style="margin-bottom: 15px;">
        <a href="${pdfUrl}" target="_blank" style="display: inline-block; padding: 12px 24px; background-color: #1a73e8; color: white; text-decoration: none; border-radius: 4px; font-weight: bold; font-size: 16px; width: 80%;">
          📄 PDFを開いて印刷する
        </a>
      </div>
      <div>
        <a href="${folderUrl}" target="_blank" style="display: inline-block; padding: 10px 20px; background-color: #f1f3f4; color: #3c4043; text-decoration: none; border-radius: 4px; font-weight: bold; font-size: 14px; border: 1px solid #dadce0; width: 80%;">
          📁 保存先のフォルダを開く
        </a>
      </div>
    </div>
  `).setWidth(450).setHeight(300);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "🎉 処理完了");
}