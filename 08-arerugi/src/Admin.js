/**
 * Admin.gs
 * 役割：メニュー表示、イベント監視、アーカイブ処理
 */

function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    ui.createMenu('🍱 アレルギー管理')
      // --- サブメニュー：初期設定 ---
      .addSubMenu(ui.createMenu('初期設定')
        .addItem('🔑 APIキーを登録・テストする', 'registerApiKeyFromSheet')
        .addItem('🚀 定期実行トリガーをまとめて一括設定する', 'registerAllTriggers')
        .addItem('❌ 設定したトリガーをすべて一括解除する', 'deleteAllTriggers')
      )
      
      .addSeparator()
      .addItem('💻 操作パネルを表示', 'showSidebar')
      .addItem('🔄 ダッシュボードを手動更新', 'refreshDashboardFormulas')
      
      .addSeparator()
      .addItem('🚨 記入漏れを自動チェック', 'runMissingEntryCheck')
      
      .addSeparator() // 区切り線
      .addItem('💬 今すぐチャットの送信', 'sendChatNow')
      
      .addSeparator()
      .addItem('📄 保護者提出用PDFを全一括作成', 'generateParentForms')

      .addSeparator()
      // --- サブメニュー：データクリア ---
      .addSubMenu(ui.createMenu('🧹 データクリア')
        .addItem('献立マスタ(Master)のみクリア', 'clearMasterDataOnly')
        .addItem('個人データ(Main)のみクリア', 'clearMainDataOnly')
        .addItem('確認用一覧(Verify)のみクリア', 'clearVerifyDataOnly')
        .addSeparator()
        .addItem('全データ一括クリア', 'clearAllData')
      )
      
      .addSeparator()
      // --- サブメニュー：メインテナンス ---
      .addSubMenu(ui.createMenu('📄 メインテナンス')
        .addItem('📋 AI引継ぎ用プロンプト作成', 'createHandoverDoc')
        .addItem('💾 全体をバックアップ', 'backupSpreadsheet')
      )
      
      .addSeparator()
      .addItem('📋 確認用一覧をPDF出力', 'exportVerifySheetAsPDF')
      
      .addSeparator()
      .addItem('📦 年度末処理（過去データをアーカイブ）', 'openArchiveDialog')
      .addToUi();
  } catch (e) {
    console.error("onOpenメニュー作成エラー: " + e.message);
  }
}

function onEdit(e) {
  if (!e) return;
  const ss = e.source;
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  const col = e.range.getColumn();
  const row = e.range.getRow();

  // 🌟 B2セル（APIキー）の編集・エンター押下を検知して自動登録
  if (sheetName === CONFIG.SHEET_NAMES.CONFIG && row === 2 && col === 2) {
    if (typeof CONFIG.getApiKey === 'function') {
      CONFIG.getApiKey();
    }
    return;
  }

  if (sheetName === CONFIG.SHEET_NAMES.DASHBOARD && row === 2 && col === 2) {
    ss.toast("日付変更を検知。Googleサイトの表示を更新中...", "⏳ 連動中");
    try {
      if (typeof refreshDashboardFormulas === 'function') refreshDashboardFormulas();
      ss.toast("最新データに更新されました。", "✅ 完了");
    } catch (err) {
      ss.toast("エラー: " + err.message, "❌ 実行失敗", 20);
      if (typeof writeLog === 'function') writeLog("onEditエラー: " + err.stack, "error");
    }
    return;
  }

  if (sheetName === CONFIG.SHEET_NAMES.MAIN && row === 1 && col === 1) {
    const val = e.range.getValue();
    const lastRow = sheet.getLastRow();
    if (lastRow >= CONFIG.START_ROW) {
      sheet.getRange(CONFIG.START_ROW, 1, lastRow - CONFIG.START_ROW + 1).setValue(val);
    }
    return;
  }

  let isTarget = false;
  if (sheetName === CONFIG.SHEET_NAMES.VERIFY && col === 8) isTarget = true;
  if (sheetName === CONFIG.SHEET_NAMES.MASTER && col === 4) isTarget = true;
  if (isTarget && row >= CONFIG.START_ROW) {
    ss.toast(`${sheetName} での修正を検知！辞書を更新します...`, "🚀 学習開始");
    if (typeof processDictionaryLearning === 'function') processDictionaryLearning(e);
  }
}

function showSidebar() {
  const html = HtmlService.createTemplateFromFile('Sidebar').evaluate().setTitle('アレルギー管理パネル').setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

// === 以下、アーカイブ（退避）機能 ===
function openArchiveDialog() {
  const html = HtmlService.createHtmlOutputFromFile('ArchiveDialog')
      .setWidth(420).setHeight(280);
  SpreadsheetApp.getUi().showModalDialog(html, '📦 過去データのアーカイブ');
}

function executeArchive(targetDateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MAIN);
  if (!mainSheet) throw new Error(`${CONFIG.SHEET_NAMES.MAIN} シートが見つかりません。`);

  const targetDate = new Date(targetDateStr.replace(/-/g, '/'));
  targetDate.setHours(23, 59, 59);

  const lastRow = mainSheet.getLastRow();
  const lastCol = mainSheet.getLastColumn();
  if (lastRow <= 1) throw new Error('アーカイブするデータがありません。');
  
  const dataRange = mainSheet.getRange(2, 1, lastRow - 1, lastCol);
  const data = dataRange.getValues();

  let archiveData = [];
  let rowsToDelete = [];

  for (let i = data.length - 1; i >= 0; i--) {
    const rowDateVal = data[i][1];
    if (rowDateVal === '毎日') continue;
    const rowDate = new Date(rowDateVal);
    if (!isNaN(rowDate.getTime()) && rowDate <= targetDate) {
      archiveData.push(data[i]);
      rowsToDelete.push(i + 2);
    }
  }

  if (archiveData.length === 0) throw new Error(`指定された日付以前のデータは見つかりませんでした。`);

  const folder = getOrCreateBackupFolder_();
  const timeStamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmm");
  const newSsName = `[アーカイブ] アレルギー対応記録_${timeStamp}`;
  const newSs = SpreadsheetApp.create(newSsName);
  
  const fileId = newSs.getId();
  const file = DriveApp.getFileById(fileId);
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);

  const newSheet = newSs.getSheets()[0];
  newSheet.setName("退避データ");
  const headers = mainSheet.getRange(1, 1, 1, lastCol).getValues();
  newSheet.getRange(1, 1, 1, lastCol).setValues(headers);
  
  archiveData.reverse();
  newSheet.getRange(2, 1, archiveData.length, lastCol).setValues(archiveData);

  rowsToDelete.forEach(row => mainSheet.deleteRow(row));

  if (typeof writeLog === 'function') writeLog(`📦 アーカイブ完了: ${archiveData.length}件退避`, "success");
  ss.toast(`✅ ${archiveData.length}件のデータを退避しました！`, 'アーカイブ完了', 8);
  return "success";
}

function getOrCreateBackupFolder_() {
  const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const file = DriveApp.getFileById(ssId);
  const parentFolders = file.getParents();
  let parentFolder = DriveApp.getRootFolder();
  if (parentFolders.hasNext()) parentFolder = parentFolders.next();

  const folderIterator = parentFolder.getFoldersByName(CONFIG.FOLDERS.BACKUP);
  if (folderIterator.hasNext()) return folderIterator.next();
  else return parentFolder.createFolder(CONFIG.FOLDERS.BACKUP);
}

/**
 * 3つのトリガーをまとめて一括登録する関数
 */
function registerAllTriggers() {
  const ui = SpreadsheetApp.getUi();
  
  // 重複登録を防ぐため、古い設定があれば先にきれいに削除する
  removeAllExistingTriggers();
  
  try {
    // 1. sendDailyNotification を 毎日【11時〜12時】に設定
    ScriptApp.newTrigger('sendDailyNotification')
      .timeBased()
      .everyDays(1)
      .atHour(11) 
      .create();

    // 2. createMonthlyVerifyList を 毎日【0時〜1時】に設定
    ScriptApp.newTrigger('createMonthlyVerifyList')
      .timeBased()
      .everyDays(1)
      .atHour(0) 
      .create();

    // 3. dailyUpdateForGoogleSites を 毎日【0時〜1時】に設定
    ScriptApp.newTrigger('dailyUpdateForGoogleSites')
      .timeBased()
      .everyDays(1)
      .atHour(0) 
      .create();
      
    // ログ記録関数（writeLog）がシステムにあれば記録する
    if (typeof writeLog === 'function') {
      writeLog("⏰ メニューから定期実行トリガーを一括設定しました（sendDailyNotification:11-12時 / createMonthlyVerifyList:0-1時 / dailyUpdateForGoogleSites:0-1時）", "success");
    }
    
    ui.alert(
      "✅ トリガーの一括設定が完了しました！\n\n" +
      "① sendDailyNotification ➔ 毎日【11:00 〜 12:00】\n" +
      "② createMonthlyVerifyList ➔ 毎日【0:00 〜 1:00】\n" +
      "③ dailyUpdateForGoogleSites ➔ 毎日【0:00 〜 1:00】\n\n" +
      "これで手動で設定し直す必要はありません。"
    );
    
  } catch (e) {
    if (typeof writeLog === 'function') writeLog(`❌ トリガー一括設定エラー: ${e.message}`, "error");
    ui.alert(`❌ トリガーの設定に失敗しました：\n${e.message}`);
  }
}

/**
 * 登録したトリガーをまとめて一括解除（削除）する関数
 */
function deleteAllTriggers() {
  const ui = SpreadsheetApp.getUi();
  const count = removeAllExistingTriggers();
  
  if (count > 0) {
    if (typeof writeLog === 'function') writeLog("🗑️ メニューからすべての定期トリガーを解除しました", "warn");
    ui.alert("✅ 設定されていた定期トリガーをすべて解除（削除）しました。");
  } else {
    ui.alert("ℹ️ 設定されている対象のトリガーは見つかりませんでした。");
  }
}

/**
 * 対象となる3つの関数の既存トリガーを削除する内部共通関数
 */
function removeAllExistingTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  const targetFunctions = ['sendDailyNotification', 'createMonthlyVerifyList', 'dailyUpdateForGoogleSites'];
  let deletedCount = 0;
  
  for (let i = 0; i < triggers.length; i++) {
    const handler = triggers[i].getHandlerFunction();
    if (targetFunctions.includes(handler)) {
      ScriptApp.deleteTrigger(triggers[i]);
      deletedCount++;
    }
  }
  return deletedCount;
}

/**
 * 🌟 メニューから「今すぐチャットの送信」が押されたときに動く関数
 */
function sendChatNow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  // 1. 押し間違い防止の確認メッセージ
  const response = ui.alert(
    "💬 チャット送信の確認", 
    "定時配信（sendDailyNotification）と同じ内容を、今すぐチャットに送信します。よろしいですか？", 
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    ss.toast("チャットの送信をキャンセルしました。", "ℹ️ キャンセル");
    return;
  }

  // 2. 送信中ステータスを画面に表示
  ss.toast("チャットへデータを送信しています。しばらくお待ちください...", "⏳ 送信中");

  try {
    
    // 🚀 定時配信の関数をここで手動実行します！
    sendDailyNotification(); 
    
    // 3. 画面右下に成功ポップアップを表示
    ss.toast("チャットへの即時送信が正常に完了しました！", "✅ 送信完了", 5);
    
  } catch (error) {
    // 万が一エラーが出た場合の安全対策
    Logger.log(error.toString());
    ui.alert("❌ 送信エラー", "チャット送信中にエラーが発生しました：\n" + error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * 🌟 メニューから「APIキーを登録・テストする」が押されたときに動く関数
 */
function registerApiKeyFromSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  try {
    const apiKey = CONFIG.getApiKey();
    if (apiKey) {
      const masked = apiKey.substring(0, 8) + "..." + apiKey.substring(apiKey.length - 4);
      ui.alert("🔑 APIキー登録・確認", `APIキーは正常に内部に登録されています。\n\n現在の登録値: ${masked}`, ui.ButtonSet.OK);
    } else {
      ui.alert("⚠️ 登録エラー", "B2セルに有効なAPIキーが入力されていないか、正しく取得できませんでした。\n「基礎データ」シートのB2セルに入力した後に実行してください。", ui.ButtonSet.OK);
    }
  } catch (e) {
    ui.alert("❌ エラー", "APIキーの登録処理中にエラーが発生しました：\n" + e.toString(), ui.ButtonSet.OK);
  }
}

/**
 * 🌟 ユーザー提供のAPIキーをスクリプトプロパティに直接設定する関数
 * スクリプトエディタからこの関数を選択して実行できます。
 */
function setMyApiKey() {
  const targetKey = "AQ.Ab8RN6JMnCf48FZH2sZ8QSTS0pr4glMKh12lw00dUfu8T6q8mA";
  PropertiesService.getScriptProperties().setProperty("GEMINI_API_KEY", targetKey);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) {
    ss.toast("APIキーを設定しました: " + targetKey.substring(0, 8) + "...", "🔑 キー設定完了");
  }
}

function checkBoundSpreadsheetUrl() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const url = ss.getUrl();
  const id = ss.getId();
  const name = ss.getName();
  Browser.msgBox("【接続テスト】\nこのスクリプトが操作しているスプレッドシート：\n名前: " + name + "\nID: " + id + "\nURL: " + url);
}