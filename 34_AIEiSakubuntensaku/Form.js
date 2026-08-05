/**
 * ファイル名: Form.gs
 * 役割: フォーム送信時に、メールアドレスから生徒を特定して「添削」シートに転記する
 * バージョン: v4.0.0
 */

function onFormSubmit(e) {
  // 1. フォームからの回答（メールアドレスと英作文）を取得
  const emailAddress = e.namedValues['メールアドレス'] ? e.namedValues['メールアドレス'][0].trim() : '';
  const englishText = e.namedValues['英作文'] ? e.namedValues['英作文'][0].trim() : '';

  // どちらかが空の場合は処理を終了
  if (!emailAddress || !englishText) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ss.getSheetByName(Config.MAIN_SHEET); 
  const rosterSheet = ss.getSheetByName(Config.ROSTER_SHEET); 

  if (!targetSheet || !rosterSheet) return;

  // 2. 「名簿」シートからメールアドレスを検索し、生徒の番号を特定
  const rosterData = rosterSheet.getDataRange().getValues();
  let studentNum = -1;

  for (let i = 1; i < rosterData.length; i++) {
    if (rosterData[i][1] === emailAddress) {
      studentNum = rosterData[i][0]; 
      break;
    }
  }

  // 名簿に登録されていないメールアドレスからの送信は無視する
  if (studentNum === -1) return;

  // 3. 「添削」シートから該当生徒の行番号を探す
  const data = targetSheet.getDataRange().getValues();
  let targetRow = -1;

  for (let i = Config.DATA_START_ROW - 1; i < data.length; i++) {
    if (data[i][0] == studentNum) { 
      targetRow = i + 1;
      break;
    }
  }

  // 4. 生徒の行が見つかったら、データを書き込む
  if (targetRow !== -1) {
    targetSheet.getRange(targetRow, Config.COL_STUDENT).setValue(englishText);
    
    // 過去の添削結果やスコアをクリアしてリセット
    targetSheet.getRange(targetRow, Config.COL_FIX, 1, 6).clearContent();
    
    // 提出日時を記録
    const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
    targetSheet.getRange(targetRow, Config.COL_TIMESTAMP).setValue(dateStr);
    
    // 進捗（〇/〇人）を更新
    if (typeof Main !== 'undefined' && Main.updateProgress) {
      Main.updateProgress();
    }
  }
}