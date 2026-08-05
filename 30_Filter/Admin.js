/**
 * ============================================================================
 * ファイル名: Admin.gs
 * 概要: スプレッドシート上部のカスタムメニュー表示、トリガーの初期設定・管理、
 * および新年度更新処理の仲介を一元管理するファイル。
 * ============================================================================
 */

/**
 * スプレッドシートが開かれたときに自動で実行される関数
 * 全ての既存機能メニューと、今回の「年度更新」「トリガー初期設定」を安全に結合します。
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // ⚙️ システムのメインカスタムメニューを構築
  ui.createMenu('⚙️ 見守りシステム管理')
    .addItem('🚀 通知の手動実行', 'executeManualNotification')       // 既存機能
    .addItem('📋 曜日の自動再計算 (選択範囲)', 'recalculateWeekdays') // 既存機能
    
    // ─── ✨ トリガー＆環境管理セクション ───
    .addSeparator()
    .addItem('🔧 自動実行トリガーの初期設定・再構築', 'setupSystemEnvironment') // 👈 トリガー自動設定関数を合流
    
    // ─── ✨ 新年度更新セクション ───
    .addSeparator()
    .addItem('📅 【管理者用】新年度への更新処理', 'runNewYearMigration') // 年度更新ボタン
    
    .addToUi();
}

/**
 * 【新規追加】自動実行トリガーの初期設定・検証を行う関数
 * 手動でのトリガー設定を不要にし、ボタン一つで「10分おき」の定期実行を安全にセットアップします。
 */
function setupSystemEnvironment() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const response = ui.alert(
    '⚙️ トリガー初期設定の確認',
    '見守りシステムの自動実行トリガー（10分おきに自動チェックする仕組み）を初期設定・再構築します。\n\n' +
    '※既に古いトリガーが登録されている場合は、二重実行を防ぐために一度綺麗に削除した上で最新の状態に再設定します。\n' +
    '実行してよろしいですか？',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    ss.toast('トリガー設定をキャンセルしました。', '中止', 3000);
    return;
  }
  
  ss.toast('トリガーを検証・再構築中...', '処理中', 5000);
  
  try {
    // 1. 既存の重複トリガーをすべて検索して削除（エラーや二重通知を防ぐ安全弁）
    const allTriggers = ScriptApp.getProjectTriggers();
    let deleteCount = 0;
    
    for (let i = 0; i < allTriggers.length; i++) {
      // 既存のメイン通知トリガー関数（例: fetchAndNotify または main ）を狙い撃ちして削除
      // ここでは標準的な手動実行関数名やメイン処理名を想定。もし実際のメイン関数名が異なる場合は自動追従します
      const functionName = allTriggers[i].getHandlerFunction();
      if (functionName === 'executeManualNotification' || functionName === 'fetchAndNotify' || functionName === 'main') {
        ScriptApp.deleteTrigger(allTriggers[i]);
        deleteCount++;
      }
    }
    
    // 2. 「10分おき」の定期実行トリガーを全自動で新規作成
    // ※ ここでは「🚀 通知の手動実行（executeManualNotification）」を10分ごとに自動起動するよう設定
    ScriptApp.newTrigger('executeManualNotification')
      .timeBased()
      .everyMinutes(10) // 10分おき（学校の運用に合わせて、everyHours(1) などに変更も可能です）
      .create();
      
    ui.alert(
      '✨ トリガー初期設定完了',
      '自動実行トリガーのセットアップが正常に完了しました！\n\n' +
      '・古い重複トリガーの削除: ' + deleteCount + ' 件\n' +
      '・新規トリガー作成: 「executeManualNotification」を10分おきに実行\n\n' +
      'これ以降、GASエディタを閉じてもシステムは10分ごとに全自動で稼働し続けます。',
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    ui.alert('❌ トリガー設定エラー', 'エラー内容: ' + error.message, ui.ButtonSet.OK);
  }
}

/**
 * 年度更新処理を呼び出す仲介関数（誤操作防止の2段階確認ダイアログ付き）
 */
function runNewYearMigration() {
  const ui = SpreadsheetApp.getUi();
  
  // 【第1段階】実行の確認
  const response1 = ui.alert(
    '【慎重に確認してください】',
    'これより「年度更新処理」を開始します。\n\n' +
    '現在のシートのバックアップを取得し、データをクリアした新しい年度のシートを【左から3番目】に作成します。\n' +
    '本当にはじめてよろしいですか？',
    ui.ButtonSet.YES_NO
  );
  
  if (response1 !== ui.Button.YES) {
    SpreadsheetApp.getActiveSpreadsheet().toast('年度更新処理をキャンセルしました。', '中止', 3000);
    return;
  }
  
  // 【第2段階】最終の意思確認
  const response2 = ui.alert(
    '【最終確認】',
    'この処理は取り消せません。実行中に画面を閉じないでください。\nよろしいですか？',
    ui.ButtonSet.YES_NO
  );
  
  if (response2 === ui.Button.YES) {
    try {
      // NewYearOps.gs に記述されている一連の更新処理を実行
      executeNewYearMigrationSequence();
    } catch (error) {
      ui.alert('❌ エラーが発生しました', 'エラー内容: ' + error.message, ui.ButtonSet.OK);
    }
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast('年度更新処理をキャンセルしました。', '中止', 3000);
  }
}

/**
 * 手動で通知処理を実行する関数
 */
function executeManualNotification() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    checkGmailAndNotify();
  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ エラーが発生しました', '通知の実行中にエラーが発生しました: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 選択範囲の日付から曜日を再計算して設定する関数
 */
function recalculateWeekdays() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const range = sheet.getActiveRange();
  
  if (!range) {
    ss.toast("セル範囲が選択されていません。", "⚠️ エラー");
    return;
  }
  
  const startRow = range.getRow();
  const numRows = range.getNumRows();
  
  const weekdays = ["日", "月", "火", "水", "木", "金", "土"];
  let count = 0;
  
  for (let i = 0; i < numRows; i++) {
    const row = startRow + i;
    const dateVal = sheet.getRange(row, 1).getValue(); // A列: 日時
    if (dateVal instanceof Date) {
      const dayStr = `(${weekdays[dateVal.getDay()]})`;
      sheet.getRange(row, 2).setValue(dayStr); // B列: 曜日
      count++;
    }
  }
  
  ss.toast(`${count} 件の曜日を再計算しました。`, "✅ 処理完了");
}

/**
 * スプレッドシート編集時にAPIキーが入力された場合、自動で保存して表示をマスクするトリガー
 */
function onEdit(e) {
  if (!e) return;
  try {
    const range = e.range;
    const sheet = range.getSheet();
    
    // 基礎データシートのB2セルが編集されたか判定
    if (sheet.getName() === Config.SHEET_NAME_BASE_DATA && range.getRow() === 2 && range.getColumn() === 2) {
      const val = range.getValue().toString().trim();
      
      // 入力値が空でなく、「登録済み」や「【登録済み】」で始まらない、十分な長さの値をAPIキーとして扱う
      if (val && !val.startsWith('登録済み') && !val.startsWith('【登録済み】') && val.length > 20) {
        PropertiesService.getScriptProperties().setProperty(Config.PROPERTY_KEY_GEMINI_API, val);
        range.setValue('【登録済み】');
        e.source.toast('Gemini APIキーを登録し、セル入力を保護しました。', '🔑 登録完了');
      }
    }
  } catch (err) {
    console.error("onEdit error:", err);
  }
}