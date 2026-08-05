/**
 * Admin.gs
 * メニュー生成・自動化トリガーの設定・バックアップ管理
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🤖 AI秘書管理')
    // --- 【1. カレンダー・タスク操作】日常のメイン業務 ---
    .addItem('🔄 ToDoをカレンダーに同期', 'admin_syncCalendar')
    .addSeparator()
    .addItem('🧹 重複ToDoの整理（クリーンアップ）', 'showDuplicateTasksDialog')
    .addItem('🙈 完了したToDoを隠す（非表示）', 'hideCompletedTasks')
    .addItem('👀 すべてのToDoを表示する', 'showAllTasks')
    .addItem('🗑️ 完了したToDoを完全に削除する', 'deleteCompletedTasks')
    .addSeparator()
    
    // --- 【2. 業務支援・文書作成】書類仕事の爆速化 ---
    .addSubMenu(ui.createMenu('📝 日誌・報告作成')
      .addItem('✨ AI日誌ドラフトを作成（自動要約）', 'generateDailyReport')
      .addItem('📋 マニュアル日誌を作成（そのまま出力）', 'createDailyReport')
      .addItem('🔔 提出物の催促を今すぐテスト', 'sendTaskReminders')
    )
    .addItem('🪄 公文書のドラフトを自動作成', 'createSchoolDocument')
    .addItem('✍️ 会議メモを議事録に清書', 'finalizeMinutes')
    .addItem('🎙️ 【AI】議事録作成を開く', 'openAiMinutesUrl') 
    .addItem('🔍 名簿照合（校閲）を一括実行', 'runRosterCheck')
    .addSeparator()
    
    // --- 【3. ショートカット】ワンクリックで確認 ---
    .addItem('🗓️ 【全体】予定カレンダーを開く', 'openMainCalendar')
    .addItem('📌 【個人】ToDoカレンダーを開く', 'openTodoCalendar')    
    .addSeparator()
    
    // --- 【4. システム管理】設定時・トラブル時のみ使用 ---
    .addSubMenu(ui.createMenu('⚙️ システム設定・自動化')
      .addItem('🚀 初期設定（全自動化セットアップ）', 'admin_setupTrigger')
      .addItem('🛑 全自動機能を停止（リセット）', 'admin_deleteTrigger')
      .addSeparator()
      .addItem('📧 メール解析を今すぐ実行（手動）', 'fetchAndFilterMail')
      .addItem('⚙️ メールの自動解析を「開始」', 'admin_setupMailTrigger')
      .addItem('📴 メールの自動解析を「停止」', 'admin_deleteMailTrigger')
      .addSeparator()
      .addItem('✅ 音声メモ連携を「ON」にする', 'turnOnVoiceMemoTrigger')
      .addItem('❌ 音声メモ連携を「OFF」にする', 'turnOffVoiceMemoTrigger')
      .addSeparator()
      .addItem('💬 チャット通知のテスト送信', 'admin_testChat')
      .addItem('📦 バックアップを作成', 'admin_createBackup')
      .addItem('ℹ️ バージョン確認', 'admin_showVersion')
    )
    .addToUi();
}

// ==========================================
// ★トリガー（自動で動く仕掛け）の構築
// ==========================================
function admin_setupTrigger() {
  const ss = SpreadsheetApp.getActive();
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    try {
      ScriptApp.deleteTrigger(t);
    } catch(e) {
      console.warn("トリガー削除失敗 (初期設定時): " + e.message);
    }
  });
  
  ScriptApp.newTrigger('admin_triggerOnOpen').forSpreadsheet(ss).onOpen().create();
  ScriptApp.newTrigger('admin_triggerOnEdit').forSpreadsheet(ss).onEdit().create();
  
  // 既存の音声メモトリガー
  ScriptApp.newTrigger('processVoiceMemo').forSpreadsheet(ss).onFormSubmit().create();
  
  // ✨【新規追加】電話伝言用のフォーム送信トリガー
  ScriptApp.newTrigger('processPhoneMemo').forSpreadsheet(ss).onFormSubmit().create();
  
  ScriptApp.newTrigger('sendDailyMorningDigest').timeBased().everyDays(1).atHour(7).create();
  ScriptApp.newTrigger('generateDailyReport').timeBased().everyDays(1).atHour(16).nearMinute(45).create();
  ScriptApp.newTrigger('sendTaskReminders').timeBased().everyDays(1).atHour(8).create();
  ScriptApp.newTrigger('fetchAndFilterMail').timeBased().everyMinutes(5).create();

  // メルアドシートの初期設定確認・作成
  ensureMailAddressSheet();

  // ToDoシートのカラム自動調整
  ensureTodoSheetColumns();

  // 各種フォルダの自動作成とURLマッピング
  ensureFoldersSetup();

  ss.toast('初期設定が完了しました！', '🤖AI秘書', 3000);
}

function admin_deleteTrigger() {
  // すべてのトリガーを一括削除するため、新機能の個別追記は不要です
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  SpreadsheetApp.getActive().toast('全ての自動実行を停止しました。', '🛑システム停止', 3000);
}

function admin_setupMailTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const status = safeTurnOnTrigger('fetchAndFilterMail', () => {
    ScriptApp.newTrigger('fetchAndFilterMail').timeBased().everyMinutes(5).create();
  });
  
  if (status === 'ALREADY_ON') {
    ss.toast('メールの自動解析はすでに「開始」されています。', '📧設定確認', 3000);
  } else if (status === 'RECREATED') {
    ss.toast('重複していたメール自動解析トリガーを整理し、再設定しました。', '📧設定完了', 3000);
  } else {
    ss.toast('メールの自動解析を開始しました。', '📧設定完了', 3000);
  }
}

function admin_deleteMailTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const deleted = safeTurnOffTrigger('fetchAndFilterMail');
  if (deleted) {
    ss.toast('メールの自動解析を停止しました。', '📴設定解除', 3000);
  } else {
    ss.toast('メールの自動解析はすでに「停止」されています。', '📴設定確認', 3000);
  }
}

function admin_syncCalendar() {
  SpreadsheetApp.getActive().toast('ToDoをカレンダーに同期中...', '📅連携', 2000);
  syncTasksToCalendar();
  SpreadsheetApp.getActive().toast('同期完了', '✅完了', 3000);
}

function admin_testChat() { sendDailyMorningDigest(); }

// ==========================================
// ★バックアップ機能（複数フォームコピークリア対応）
// ==========================================
function createBackupCore(ss) {
  const parentFolder = DriveApp.getFileById(ss.getId()).getParents().next();
  let backupFolder;
  const folders = parentFolder.getFoldersByName(Config.BACKUP_FOLDER_NAME);
  if (folders.hasNext()) backupFolder = folders.next();
  else backupFolder = parentFolder.createFolder(Config.BACKUP_FOLDER_NAME);
  
  const fileName = `[${Config.VERSION}]_${Utilities.formatDate(new Date(), "JST", "yyyyMMdd_HHmmss")}_${ss.getName()}`;
  const copiedFile = DriveApp.getFileById(ss.getId()).makeCopy(fileName, backupFolder);
  
  try {
    const copiedSs = SpreadsheetApp.openById(copiedFile.getId());
    const sheets = copiedSs.getSheets();
    
    sheets.forEach(sheet => {
      const formUrl = sheet.getFormUrl();
      if (formUrl) {
        try {
          const formId = FormApp.openByUrl(formUrl).getId();
          DriveApp.getFileById(formId).setTrashed(true);
        } catch(err) {
          console.warn("フォーム削除スキップ: " + err);
        }
      }
    });
  } catch(e) {
    console.error("フォームコピー一括削除エラー: " + e);
  }
}

function admin_createBackup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('バックアップを作成しています...', '📦 処理中', 3000);
  createBackupCore(ss);
  ss.toast('バックアップを作成し、不要なフォームのコピーを自動削除しました。', '📦 完了', 4000);
}

function admin_showVersion() {
  SpreadsheetApp.getUi().alert(`現在のバージョン: ${Config.VERSION}\n対象年度: ${Config.FISCAL_YEAR}`);
}

// ==========================================
// ★ショートカット・その他管理
// ==========================================
function openMainCalendar() { openCalendarByCell_(Config.MASTER_POS.CALENDAR_ID_CELL, '全体の予定カレンダー'); }
function openTodoCalendar() { openCalendarByCell_(Config.MASTER_POS.TODO_CALENDAR_ID_CELL, 'ToDo専用カレンダー'); }

function openCalendarByCell_(cellPos, title) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(Config.SHEET_NAME_MASTER);
  const calId = masterSheet.getRange(cellPos).getValue();
  if (!calId) { SpreadsheetApp.getUi().alert(`基礎データシートの ${cellPos} セルにカレンダーIDがありません。`); return; }
  const calUrl = `https://calendar.google.com/calendar/u/0/r?cid=${encodeURIComponent(calId)}`;
  const html = HtmlService.createHtmlOutput(`<html><body style="font-family: sans-serif; text-align: center; padding: 20px;"><h3>📅 ${title}</h3><a href="${calUrl}" target="_blank" onclick="google.script.host.close();" style="padding: 10px 20px; background-color: #1a73e8; color: white; text-decoration: none; border-radius: 4px;">カレンダーを開く</a></body></html>`).setWidth(320).setHeight(150);
  SpreadsheetApp.getUi().showModelessDialog(html, `ショートカット`);
}

function openAiMinutesUrl() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(Config.SHEET_NAME_MASTER);
  if (!masterSheet) { SpreadsheetApp.getUi().alert(`エラー: 基礎データシートが見つかりません。`); return; }
  
  const url = masterSheet.getRange('B20').getValue();
  if (!url || !url.toString().startsWith('http')) {
    SpreadsheetApp.getUi().alert(`基礎データシートの B20 セルに正しいURLが入力されていません。`);
    return;
  }
  const html = HtmlService.createHtmlOutput(`
    <html>
    <body style="font-family: sans-serif; text-align: center; padding: 15px; margin: 0; background-color: #f8f9fa;">
      <h3 style="color: #1a73e8; margin-bottom: 5px;">🎙️ AI議事録作成アプリ</h3>
      <p style="font-size: 13px; color: #5f6368; margin-top: 0;">下のボタンを押すと文字起こし画面が開きます。</p>
      <a href="${url}" target="_blank" onclick="google.script.host.close();" 
         style="display: inline-block; padding: 12px 24px; background-color: #1a73e8; color: white; text-decoration: none; border-radius: 4px; font-weight: bold;">
        アプリを起動する
      </a>
    </body>
    </html>
  `).setWidth(320).setHeight(140);
  SpreadsheetApp.getUi().showModelessDialog(html, `ショートカット`);
}

// ==========================================
// ★ダッシュボードの全自動同期・編集検知
// ==========================================
function admin_triggerOnOpen() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(Config.SHEET_NAME_DASHBOARD);
  if (!sheet) return;
  const today = new Date();
  sheet.getRange('B1').setValue(today);
  admin_updateDashboardByDate(today, sheet);
}

function admin_triggerOnEdit(e) {
  try {
    if (!e) return;
    const range = e.range;
    const sheet = range.getSheet();
    const sheetName = sheet.getName();
  
    // 1. 基礎データシートの編集検知（APIキーの格納等）
    if (sheetName === Config.SHEET_NAME_MASTER) {
      if (range.getA1Notation() === Config.MASTER_POS.API_KEY_CELL) {
        const val = range.getValue().toString().trim();
        const skipValues = ["（設定済み）", "（格納済み）", "１本格納しました", "1本格納しました"];
        const isMasked = skipValues.some(v => val.startsWith(v));
        
        if (val !== "" && !isMasked) {
          // カンマ、改行、スペース等で分割してAPIキーを抽出
          const keys = val.split(/[\s,，、\n\r]+/).filter(k => k.length > 20);
          
          if (keys.length > 0) {
            const props = PropertiesService.getScriptProperties();
            props.setProperty('GEMINI_API_KEYS', JSON.stringify(keys));
            props.setProperty('GEMINI_API_KEY', keys[0]); // 互換性のため最初のキーを保存
            
            if (keys.length > 1) {
              range.setValue(`（設定済み） (${keys.length}本のキー)`);
              e.source.toast(`APIキーを${keys.length}本登録し、保護しました。`, '🔑 複数キー登録完了');
            } else {
              range.setValue("（設定済み）");
              e.source.toast("APIキーを安全に格納しました。", "🔑 格納完了");
            }
            SpreadsheetApp.flush();
          } else {
            e.source.toast("有効なAPIキー（20文字以上）が見つかりませんでした。", "⚠️ 登録エラー");
          }
        } else if (val === "") {
          const props = PropertiesService.getScriptProperties();
          props.deleteProperty('GEMINI_API_KEY');
          props.deleteProperty('GEMINI_API_KEYS');
          e.source.toast("APIキーを削除しました。", "🗑️ 削除完了");
        }
      }
      return;
    }
  
    // 2. ダッシュボードの編集検知
    if (sheetName === Config.SHEET_NAME_DASHBOARD) {
      if (range.getRow() === 1 && range.getColumn() === 2) {
        const inputValue = range.getValue();
        if (inputValue instanceof Date || !isNaN(Date.parse(inputValue))) {
          const targetDate = new Date(inputValue);
          e.source.toast('日付変更を検知。行事予定を自動取得します...', '🔄 AI自動同期', 3000);
          admin_updateDashboardByDate(targetDate, sheet);
        }
      }
    }
  
    // 3. ToDoシートの編集検知（転送先・氏名の同期、およびアクション実行）
    if (sheetName === Config.SHEET_NAME_TODO) {
      const rowIdx = range.getRow();
      if (rowIdx < 2) return; // 🚨ヘッダー行（1行目）の編集は無視する
      
      const actionCol = Config.TODO_COL.ACTION + 1;
      const stakeholderCol = Config.TODO_COL.STAKEHOLDER + 1;
      const nameCol = Config.TODO_COL.NAME + 1;
      
      // a. アクションチェックボックスの検知
      if (range.getColumn() === actionCol && range.getValue() === true) {
        const rowIdx = range.getRow();
        if (rowIdx >= 2) {
          // 多重実行を防ぐために即座にチェックボックスをクリア（falseに戻す）
          range.setValue(false);
          SpreadsheetApp.flush();
          
          // 転送アクションの実行
          handleTodoAction(sheet, rowIdx);
        }
        return;
      }
      
      // 複数選択（追加・削除・トグル）の処理
      let finalValue = range.getValue().toString().trim();
      
      if ((range.getColumn() === stakeholderCol || range.getColumn() === nameCol) && e && e.oldValue) {
        const oldValue = e.oldValue.toString().trim();
        const newValue = e.value ? e.value.toString().trim() : "";
        
        // 古い値と新しい値が異なり、かつ新しい値が空でなければトグル処理を行う
        if (oldValue && newValue && oldValue !== newValue && !oldValue.startsWith("（設定済み）")) {
          // カンマ、読点、改行等で分割 (スペースでの分割はフルネーム内のスペースを誤分割するため除外)
          let list = oldValue.split(/[,，、\n\r]+/).filter(v => v.trim() !== "");
          const index = list.indexOf(newValue);
          
          if (index > -1) {
            // すでに含まれている場合は、選択解除（リストから削除）
            list.splice(index, 1);
          } else {
            // 含まれていない場合は、追加
            if (newValue === '手元で留める（転送不要）') {
              list = [newValue];
            } else {
              // 「手元で留める（転送不要）」があれば除外して追加
              list = list.filter(v => v !== '手元で留める（転送不要）');
              list.push(newValue);
            }
          }
          
          finalValue = list.join(", ");
          range.setValue(finalValue);
          SpreadsheetApp.flush();
        }
      }
      
      // b. 転送先（役職・係等）の編集検知 -> 氏名を自動更新
      if (range.getColumn() === stakeholderCol) {
        const rowIdx = range.getRow();
        const nameRange = sheet.getRange(rowIdx, nameCol);
        
        if (finalValue === '手元で留める（転送不要）' || finalValue === '') {
          nameRange.setValue('');
        } else {
          const correspondingName = getNameByRole(finalValue);
          if (correspondingName) {
            nameRange.setValue(correspondingName);
          }
        }
        return;
      }
      
      // c. 氏名の編集検知 -> 役職・係等を自動更新
      if (range.getColumn() === nameCol) {
        const rowIdx = range.getRow();
        const stakeholderRange = sheet.getRange(rowIdx, stakeholderCol);
        
        if (finalValue === '') {
          // 何もしない
        } else {
          const correspondingRole = getRoleByName(finalValue);
          if (correspondingRole) {
            stakeholderRange.setValue(correspondingRole);
          }
        }
        return;
      }
    }
  } catch(err) {
    console.error("admin_triggerOnEdit error:", err);
    SpreadsheetApp.getActiveSpreadsheet().toast("編集処理中にエラーが発生しました: " + err.toString(), "⚠️ エラー発生", 6000);
  }
}

function admin_updateDashboardByDate(targetDate, sheet) {
  const dayLabels = ['(日)', '(月)', '(火)', '(水)', '(木)', '(金)', '(土)'];
  sheet.getRange('C1').setValue(dayLabels[targetDate.getDay()]);
  
  try {
    const masterSheet = sheet.getParent().getSheetByName(Config.SHEET_NAME_MASTER);
    if (!masterSheet) return;
    
    // 【修正】全体用（B6）とToDo用（B18）の両方のカレンダーIDを取得
    const calId = masterSheet.getRange(Config.MASTER_POS.CALENDAR_ID_CELL).getValue();
    const todoCalId = masterSheet.getRange(Config.MASTER_POS.TODO_CALENDAR_ID_CELL).getValue();
    
    const calendar = CalendarApp.getCalendarById(calId);
    const todoCalendar = CalendarApp.getCalendarById(todoCalId);

    const startTime = new Date(targetDate.getFullYear(), targetDate.getMonth(), targetDate.getDate(), 0, 0, 0);
    const endTime = new Date(targetDate.getFullYear(), targetDate.getMonth(), targetDate.getDate(), 23, 59, 59);
    
    // 【修正】両方のカレンダーから今日の予定を取得して1つに合体させる
    let allEvents = [];
    if (calendar) {
      allEvents = allEvents.concat(calendar.getEvents(startTime, endTime));
    }
    if (todoCalendar) {
      allEvents = allEvents.concat(todoCalendar.getEvents(startTime, endTime));
    }
    
    let memoRow = Math.max(sheet.getLastRow(), 20);
    const values = sheet.getRange(1, 1, memoRow, 1).getValues();
    for (let i = 0; i < values.length; i++) {
      if (values[i][0] && values[i][0].toString().includes('一般メモ・記録')) { memoRow = i + 1; break; }
    }
    
    const endClearRow = memoRow > 4 ? memoRow - 1 : 20;
    if (endClearRow >= 4) { sheet.getRange(4, 1, endClearRow - 4 + 1, 6).clearContent().clearDataValidations(); }
    
    const rows = [];
    allEvents.forEach(event => {
      let startStr = event.isAllDayEvent() ? '終日' : Utilities.formatDate(event.getStartTime(), 'JST', 'HH:mm');
      let endStr = event.isAllDayEvent() ? '' : Utilities.formatDate(event.getEndTime(), 'JST', 'HH:mm');
      rows.push([false, startStr, endStr, event.getTitle(), event.getDescription(), event.getId()]);
    });
    
    if (rows.length > 0) {
      const kyushokuIcons = ['🍴', '🍽', '🥛', '🍞', '🍛', '🍚', '🍱', '🍙', '🍲', '🥗', '🥄', '🥖', '🍜'];
      rows.sort((a, b) => {
        const titleA = a[3] ? a[3].toString().trim() : '';
        const titleB = b[3] ? b[3].toString().trim() : '';
        const isKyushokuA = kyushokuIcons.some(icon => titleA.startsWith(icon)) || titleA.includes('給食') || titleA.includes('献立');
        const isKyushokuB = kyushokuIcons.some(icon => titleB.startsWith(icon)) || titleB.includes('給食') || titleB.includes('献立');
        if (isKyushokuA && !isKyushokuB) return -1;
        if (!isKyushokuA && isKyushokuB) return 1;
        return 0;
      });

      const availableRows = memoRow > 4 ? (memoRow - 4) : 15;
      if (rows.length > availableRows && memoRow > 0) { sheet.insertRowsBefore(memoRow, rows.length - availableRows); }
      sheet.getRange(4, 1, rows.length, 6).setValues(rows);
      sheet.getRange(4, 1, rows.length, 1).insertCheckboxes();
    }
    // 【修正】通知メッセージを変更
    sheet.getParent().toast(`${Utilities.formatDate(targetDate, 'JST', 'M/d')} の【行事予定とToDo】を表示しました。`, '✅ 同期完了', 2000);
  } catch (e) { console.error(e); }
}

function addScheduleToCalendar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(Config.SHEET_NAME_DASHBOARD);
  if (!sheet) return;

  const b1Value = sheet.getRange('B1').getValue();
  let targetDate = new Date(b1Value);
  if (!b1Value || isNaN(targetDate.getTime())) {
    ss.toast('B1セルの日付が正しくありません。', '❌登録エラー', 3000); return;
  }

  const masterSheet = ss.getSheetByName(Config.SHEET_NAME_MASTER);
  const calId = masterSheet.getRange(Config.MASTER_POS.CALENDAR_ID_CELL).getValue();
  const todoCalId = masterSheet.getRange(Config.MASTER_POS.TODO_CALENDAR_ID_CELL).getValue();
  
  const calendar = CalendarApp.getCalendarById(calId);
  const todoCalendar = CalendarApp.getCalendarById(todoCalId);
  if (!calendar) { ss.toast('カレンダーの取得に失敗しました。', '❌登録エラー', 3000); return; }

  let memoRow = 0;
  const maxRow = Math.max(sheet.getLastRow(), 30);
  // 【修正】F列（6列目）の既存カレンダーイベントIDまで読み込むように「5」から「6」に変更
  const values = sheet.getRange(1, 1, maxRow, 6).getValues(); 
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] && values[i][0].toString().includes('一般メモ・記録')) { memoRow = i + 1; break; }
  }
  if (memoRow === 0) memoRow = 21;

  let addedCount = 0;
  for (let rowIdx = 3; rowIdx < memoRow - 1; rowIdx++) {
    const startTimeStr = values[rowIdx][1] ? values[rowIdx][1].toString().trim() : '';
    const endTimeStr = values[rowIdx][2] ? values[rowIdx][2].toString().trim() : '';
    const title = values[rowIdx][3] ? values[rowIdx][3].toString().trim() : '';
    const description = values[rowIdx][4] ? values[rowIdx][4].toString().trim() : '';
    const eventId = values[rowIdx][5] ? values[rowIdx][5].toString().trim() : ''; // 【追加】カレンダー上のID

    // 【重要】すでにカレンダーIDがある（画面に自動表示された）予定は二重登録を防ぐため絶対にスキップ
    if (eventId) continue;

    if (title) {
      // 【修正】手書きされたタイトルに「📌」や「ToDo」があれば個人用、なければ全体用に自動振り分け
      let targetCalendar = calendar;
      if ((title.includes('📌') || title.includes('ToDo')) && todoCalendar) {
        targetCalendar = todoCalendar;
      }

      if (startTimeStr === '終日' || !startTimeStr) {
        targetCalendar.createAllDayEvent(title, targetDate, {description: description});
      } else {
        try {
          const startParts = startTimeStr.split(':');
          const endParts = endTimeStr.split(':');
          const startDateTime = new Date(targetDate.getFullYear(), targetDate.getMonth(), targetDate.getDate(), parseInt(startParts[0], 10), parseInt(startParts[1], 10), 0);
          let endDateTime;
          if (endParts.length === 2) {
            endDateTime = new Date(targetDate.getFullYear(), targetDate.getMonth(), targetDate.getDate(), parseInt(endParts[0], 10), parseInt(endParts[1], 10), 0);
          } else {
            endDateTime = new Date(startDateTime.getTime() + 60 * 60 * 1000);
          }
          targetCalendar.createEvent(title, startDateTime, endDateTime, {description: description});
        } catch(timeError) {
          targetCalendar.createAllDayEvent(title, targetDate, {description: description});
        }
      }
      addedCount++;
    }
  }

  if (addedCount > 0) {
    ss.toast(`${addedCount} 件の新規予定をカレンダーに登録しました！`, '📤 登録完了', 4000);
    admin_updateDashboardByDate(targetDate, sheet);
  } else {
    ss.toast('新しく書き足された予定は見つかりませんでした。', 'ℹ️ 通知', 3000);
  }
}

function deleteSelectedSchedules() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(Config.SHEET_NAME_DASHBOARD);
  if (!sheet) return;

  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('⚠️ 行事削除の確認', 'チェックを入れた【学校行事】を全体カレンダーから完全に削除しますか？\n（※教頭先生の個人ToDoは安全のため、ここでは絶対に削除されません）', ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) { ss.toast('削除をキャンセルしました。', 'ℹ️ キャンセル', 3000); return; }

  const masterSheet = ss.getSheetByName(Config.SHEET_NAME_MASTER);
  const calId = masterSheet.getRange(Config.MASTER_POS.CALENDAR_ID_CELL).getValue();
  const calendar = CalendarApp.getCalendarById(calId);

  let memoRow = Math.max(sheet.getLastRow(), 20);
  const values = sheet.getRange(1, 1, memoRow, 6).getValues(); 
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] && values[i][0].toString().includes('一般メモ・記録')) { memoRow = i + 1; break; }
  }

  let deletedCount = 0;
  let unassignedCount = 0;
  let skippedTodoCount = 0;

  for (let rowIdx = memoRow - 2; rowIdx >= 3; rowIdx--) {
    const isChecked = values[rowIdx][0];
    const eventId = values[rowIdx][5]; 
    const title = values[rowIdx][3] ? values[rowIdx][3].toString() : '';

    if (isChecked === true) {
      // 📌 や ToDo という文字が入っている行は、個人ToDoなので絶対に消さずにスキップする（安全ガード）
      if (title.includes('📌') || title.includes('ToDo')) {
        skippedTodoCount++;
        // チェックを自動で外す
        sheet.getRange(rowIdx + 1, 1).setValue(false);
        continue; 
      }

      if (eventId && calendar) {
        try {
          const event = calendar.getEventById(eventId);
          if (event) { event.deleteEvent(); deletedCount++; }
        } catch(e) { console.error(e); }
      } else if (!eventId) {
        // カレンダーに登録されていない、ダッシュボード上の手書きの空行なら削除する
        sheet.deleteRow(rowIdx + 1);
        unassignedCount++;
      }
    }
  }

  let msg = `全体カレンダーから行事を ${deletedCount} 件削除しました。`;
  if (skippedTodoCount > 0) {
    msg += `（※個人ToDo ${skippedTodoCount} 件は安全のためスキップしました）`;
  }
  ss.toast(msg, '🗑️ 処理完了', 5000);

  const targetDate = new Date(sheet.getRange('B1').getValue());
  admin_updateDashboardByDate(targetDate, sheet); 
}

// ==========================================
// ★ トリガー安全管理ヘルパー
// ==========================================

/**
 * 特定の関数に対するトリガーを安全にON（有効化）にする。
 * 重複がある場合はクリーンアップして1つにする。既に1つだけある場合はスキップする。
 * @param {string} functionName 対象の関数名
 * @param {Function} createTriggerCallback トリガーを作成するコールバック関数
 * @return {string} 状態（'ALREADY_ON', 'CREATED', 'RECREATED'）
 */
function safeTurnOnTrigger(functionName, createTriggerCallback) {
  const triggers = ScriptApp.getProjectTriggers();
  const matchedTriggers = triggers.filter(t => t.getHandlerFunction() === functionName);
  
  if (matchedTriggers.length === 1) {
    return 'ALREADY_ON';
  }
  
  if (matchedTriggers.length > 1) {
    matchedTriggers.forEach(t => {
      try {
        ScriptApp.deleteTrigger(t);
      } catch(e) {
        console.error(`トリガー削除エラー (${functionName}): ${e}`);
      }
    });
    createTriggerCallback();
    return 'RECREATED';
  }
  
  createTriggerCallback();
  return 'CREATED';
}

/**
 * 特定の関数に対するトリガーを安全にOFF（無効化）にする。
 * @param {string} functionName 対象の関数名
 * @return {boolean} 削除されたトリガーがあったかどうか
 */
function safeTurnOffTrigger(functionName) {
  const triggers = ScriptApp.getProjectTriggers();
  let deleted = false;
  triggers.forEach(t => {
    if (t.getHandlerFunction() === functionName) {
      try {
        ScriptApp.deleteTrigger(t);
        deleted = true;
      } catch(e) {
        console.error(`トリガー削除エラー (${functionName}): ${e}`);
      }
    }
  });
  return deleted;
}