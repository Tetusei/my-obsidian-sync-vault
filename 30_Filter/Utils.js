/**
 * @file Utils.gs
 * @description 各種データ取得、セキュリティ管理、バックアップ、タイマー設定などの共通機能を管理します。
 */

function getWebhookUrl() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(Config.SHEET_NAME_BASE_DATA);
  if (!sheet) throw new Error(`「${Config.SHEET_NAME_BASE_DATA}」シートが見つかりません。`);
  return sheet.getRange(Config.WEBHOOK_CELL).getValue().toString().trim();
}

/**
 * 基礎データシートのB3セルから対象者フィルターの文字列を取得する
 */
function getTargetPerson() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(Config.SHEET_NAME_BASE_DATA);
  if (!sheet) throw new Error(`「${Config.SHEET_NAME_BASE_DATA}」シートが見つかりません。`);
  return sheet.getRange(Config.TARGET_PERSON_CELL).getValue().toString().trim();
}

/**
 * 安全な内部保管庫からGemini APIキーを安全に呼び出す
 */
function getStoredApiKey() {
  const apiKey = PropertiesService.getScriptProperties().getProperty(Config.PROPERTY_KEY_GEMINI_API);
  if (!apiKey) {
    throw new Error('Gemini APIキーが設定されていません。「基礎データ」シートのB2セルに入力してエンターキーを押してください。');
  }
  return apiKey;
}

function getStudentDictionary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(Config.SHEET_NAME_ROSTER);
  if (!sheet) throw new Error(`「${Config.SHEET_NAME_ROSTER}」シートが見つかりません。`);
  
  const data = sheet.getDataRange().getValues();
  const dict = {};
  for (let i = 1; i < data.length; i++) {
    const email = data[i][0]; // A列: メールアドレス
    if (email) {
      dict[email.toString().trim()] = {
        grade: data[i][1],  // B列: 学年
        class: data[i][2],  // C列: 組
        number: data[i][3], // D列: 番号
        name: data[i][4]    // E列: 氏名
      };
    }
  }
  return dict;
}

function createBackup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast("バックアップを作成中です...", "処理中");
  const file = DriveApp.getFileById(ss.getId());
  const parentFolder = file.getParents().next(); 
  let backupFolder;
  const folders = parentFolder.getFoldersByName(Config.BACKUP_FOLDER_NAME);
  if (folders.hasNext()) {
    backupFolder = folders.next();
  } else {
    backupFolder = parentFolder.createFolder(Config.BACKUP_FOLDER_NAME);
  }
  const dateStr = Utilities.formatDate(new Date(), "JST", "yyyyMMdd_HHmm");
  const backupName = `[${Config.VERSION}]_${dateStr}_${ss.getName()}`;
  file.makeCopy(backupName, backupFolder);
  ss.toast(`バックアップ完了: ${backupName}`, "✅ 成功", 5);
}

function setAutomaticTrigger() {
  const ui = SpreadsheetApp.getUi();
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === Config.TRIGGER_FUNCTION_MAIL) {
      ui.alert('お知らせ', 'すでに自動メール確認（10分おき）がオンになっています。', ui.ButtonSet.OK);
      return;
    }
  }
  ScriptApp.newTrigger(Config.TRIGGER_FUNCTION_MAIL).timeBased().everyMinutes(Config.TRIGGER_INTERVAL_MINUTES).create();
  ui.alert('設定完了', `自動メール確認をオンにしました。\n${Config.TRIGGER_INTERVAL_MINUTES}分おきに自動確認し、AI解析とシート自動追記を行います。`, ui.ButtonSet.OK);
}

function removeAutomaticTrigger() {
  const ui = SpreadsheetApp.getUi();
  const triggers = ScriptApp.getProjectTriggers();
  let deleted = false;
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === Config.TRIGGER_FUNCTION_MAIL) {
      ScriptApp.deleteTrigger(trigger);
      deleted = true;
    }
  }
  const msg = deleted ? '自動メール確認をオフにしました。' : '自動メール確認はすでにオフになっています。';
  ui.alert('設定完了', msg, ui.ButtonSet.OK);
}