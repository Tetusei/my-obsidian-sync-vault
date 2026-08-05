/**
 * ファイル名: FormManager.gs
 * 役割: フォーム自動作成・締め切り・履歴管理・自動分析
 * バージョン: v5.3.0 (トースト表示自動消去対応)
 */

const FormManager = {
  createStudentForm: function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    
    const questionPrompt = ui.prompt(
      '📝 新しい問題の設定',
      '新しい英作文の問題（テーマ）を入力してください。\n（※ここで入力した文が、新しいフォームと「添削」シートのB1セルにセットされます。古い問題は履歴シートに安全に保存されます）\n\n※後で直接入力する場合は、空欄のまま「OK」を押してください。',
      ui.ButtonSet.OK_CANCEL
    );
    if (questionPrompt.getSelectedButton() !== ui.Button.OK) return;
    const newQuestionText = questionPrompt.getResponseText(); 

    const response = ui.alert('フォーム自動作成', '新しい提出フォームを作成し、「添削」シートのデータをリセットしますか？\n（※現在の状態は「履歴シート」として右側に保存されます）', ui.ButtonSet.YES_NO);
    if (response !== ui.Button.YES) return;

    ss.toast('現在のデータをバックアップしています...', '📦 保存中', -1);
    DataSync.createBackup(); 
    Utilities.sleep(1500); 

    ss.toast('履歴シートを作成し、新しいフォームを準備しています...', '🛠️ 作成中', -1);
    
    try {
      const oldTriggers = ScriptApp.getProjectTriggers();
      for (let i = 0; i < oldTriggers.length; i++) {
        if (oldTriggers[i].getHandlerFunction() === 'autoCloseForm') ScriptApp.deleteTrigger(oldTriggers[i]);
      }

      const targetSheet = ss.getSheetByName(Config.MAIN_SHEET);
      if (targetSheet) {
        
        const allSheetsForCount = ss.getSheets();
        let maxCount = 0;
        const regex = /^第(\d+)回/; 
        
        for (let i = 0; i < allSheetsForCount.length; i++) {
          const match = allSheetsForCount[i].getName().match(regex);
          if (match) {
            const count = parseInt(match[1], 10);
            if (count > maxCount) maxCount = count;
          }
        }
        const nextCount = maxCount + 1; 

        const historyDateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd HH:mm');
        const historySheetName = `第${nextCount}回 (${historyDateStr})`;
        
        let copiedSheet = null; 
        
        try {
          copiedSheet = targetSheet.copyTo(ss);
          copiedSheet.setName(historySheetName);
          copiedSheet.setTabColor('#cccccc'); 
          
          ss.setActiveSheet(copiedSheet);
          ss.moveActiveSheet(ss.getNumSheets()); 
        } catch (e) {
          console.warn('履歴シートの作成をスキップしました: ' + e.message);
        }

        if (copiedSheet) {
          ss.toast('過去のデータをAIが分析しています...', '📊 分析中', -1);
          const lastRow = copiedSheet.getLastRow();
          if (lastRow >= Config.DATA_START_ROW) {
            const texts = copiedSheet.getRange(Config.DATA_START_ROW, Config.COL_STUDENT, lastRow - Config.DATA_START_ROW + 1, 1).getValues().flat().filter(t => t);
            
            if (texts.length > 0) {
              try {
                const apiKeys = getApiKeys();
                const prompt = `あなたは優秀な中学校の英語教師です。以下の生徒の英作文を分析し、レポートを作成してください。\n\n# 📊 クラス全体の傾向と分析レポート\n## 1. よくある間違いと傾向\n## 2. 次回の授業で取り上げるべき文法ポイント\n## 3. 先生への指導アドバイス\n\n【リスト】\n${texts.join('\n')}`;
                const payload = { "contents": [{ "parts": [{ "text": prompt }] }] };
                const options = { "method": "post", "contentType": "application/json", "payload": JSON.stringify(payload), "muteHttpExceptions": true };
                
                let reportText = null;
                const maxRetries = Math.max(3, apiKeys.length * 2);
                for (let i = 0; i < maxRetries; i++) {
                  const currentKey = apiKeys[i % apiKeys.length];
                  const url = `https://generativelanguage.googleapis.com/v1beta/models/${Config.MODEL_NAME}:generateContent?key=${currentKey}`;
                  const res = UrlFetchApp.fetch(url, options);
                  if (res.getResponseCode() === 429) { Utilities.sleep(2000); continue; }
                  const json = JSON.parse(res.getContentText());
                  if (json.candidates) { reportText = json.candidates[0].content.parts[0].text; break; }
                }
                
                if (reportText) {
                  copiedSheet.getRange('B44').setValue('📊 AI分析レポート').setBackground('#e1f5fe').setFontWeight('bold');
                  copiedSheet.getRange('B45').setValue(reportText).setWrap(true);
                  copiedSheet.setRowHeight(45, 150); 
                }
              } catch(e) {
                console.warn('分析レポートの自動作成に失敗しました: ' + e.message);
              }
            }
          }
        }

        ss.setActiveSheet(targetSheet);

        const maxRowTarget = targetSheet.getMaxRows();
        if (maxRowTarget >= Config.DATA_START_ROW) {
          const numRowsToClear = maxRowTarget - Config.DATA_START_ROW + 1;
          const numColsToClear = Config.COL_MAIL_STATUS - Config.COL_STUDENT + 1;
          targetSheet.getRange(Config.DATA_START_ROW, Config.COL_STUDENT, numRowsToClear, numColsToClear).clearContent();
        }

        targetSheet.getRange(Config.QUESTION_CELL).setValue(newQuestionText);

        const deadlineRange = targetSheet.getRange(Config.DEADLINE_CELL);
        deadlineRange.clearContent();
        deadlineRange.setNumberFormat('yyyy/MM/dd'); 
        deadlineRange.setNote('【締め切り設定】\nダブルクリックしてカレンダーから日付を選び、メニューから設定を実行してください。\n（※時間は自動的に当日の 22:00 に設定されます）');
      }

      const oldSheets = ss.getSheets();
      for (let i = 0; i < oldSheets.length; i++) {
        const sheet = oldSheets[i];
        const sheetName = sheet.getName();
        if (sheetName.indexOf('フォームの回答') !== -1 || sheetName.indexOf('回答データ') !== -1) {
          const formUrl = sheet.getFormUrl();
          if (formUrl) {
            try { FormApp.openByUrl(formUrl).removeDestination(); } catch (e) {}
          }
          if (ss.getSheets().length > 1) ss.deleteSheet(sheet);
        }
      }

      const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd HH:mm');
      const form = FormApp.create(`AI英作文 提出フォーム (${dateStr})`);
      
      let formDesc = '書いた英作文を入力して送信してください。\n※あなたのアカウント情報は自動で記録されるため、名前や番号の入力は不要です。';
      if (newQuestionText) {
        formDesc = `【本日の問題】\n${newQuestionText}\n\n------------------------\n${formDesc}`;
      }
      form.setDescription(formDesc);
      form.setCollectEmail(true);
      
      const formFile = DriveApp.getFileById(form.getId());
      const parents = DriveApp.getFileById(ss.getId()).getParents();
      if (parents.hasNext()) {
        const parentFolder = parents.next();
        const folders = parentFolder.getFoldersByName('📝_提出フォーム一覧');
        formFile.moveTo(folders.hasNext() ? folders.next() : parentFolder.createFolder('📝_提出フォーム一覧'));
      }
      const textItem = form.addParagraphTextItem();
      textItem.setTitle('英作文');
      textItem.setHelpText('英語で文章を入力してください。');
      textItem.setRequired(true);
      
      form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
      
      const triggers = ScriptApp.getProjectTriggers();
      for (let i = 0; i < triggers.length; i++) {
        if (triggers[i].getHandlerFunction() === 'onFormSubmit') ScriptApp.deleteTrigger(triggers[i]);
      }
      ScriptApp.newTrigger('onFormSubmit').forSpreadsheet(ss).onFormSubmit().create();
      
      SpreadsheetApp.flush();
      Utilities.sleep(3000); 
      SpreadsheetApp.flush();

      const allSheetsFinal = ss.getSheets();
      for (let i = 0; i < allSheetsFinal.length; i++) {
        if (allSheetsFinal[i].getName().indexOf('フォームの回答') !== -1) {
          ss.setActiveSheet(allSheetsFinal[i]);
          ss.moveActiveSheet(ss.getNumSheets()); 
          allSheetsFinal[i].setName('📥_回答データ'); 
          break;
        }
      }

      if (targetSheet) {
        ss.setActiveSheet(targetSheet);
      }
      
      if (typeof Main !== 'undefined' && Main.updateProgress) Main.updateProgress();

      // ★ 追加: 出っぱなしのトーストを「完了（5秒で自動消去）」で上書きして消す
      ss.toast('フォームの準備がすべて完了しました！', '✨ 完了', 5);

      const htmlContent = `
        <div style="font-family: sans-serif; padding: 10px; color: #333;">
          <p>生徒用フォームの作成が完了し、シートをリセットしました！</p>
          <div style="margin-top: 15px; padding: 12px; background-color: #fff3cd; border: 1px solid #ffe69c; border-radius: 5px; color: #856404;">
            <strong>⚠️【重要：最後の手動設定をお願いします】</strong><br>
            <strong>「✏️ 編集用フォームを開く」</strong>から設定画面を開き、<br>
            <strong>「設定」＞「回答」＞「メールアドレスを収集する」を『確認済み』</strong>に変更してください！
          </div>
          <div style="margin-top: 20px;"><strong>【先生用URL（編集用）】</strong><br><a href="${form.getEditUrl()}" target="_blank" style="color: #1a73e8; text-decoration: none;">✏️ 編集用フォームを開く</a></div>
          <div style="margin-top: 15px;"><strong>【生徒に配るURL（回答用）】</strong><br><a href="${form.getPublishedUrl()}" target="_blank" style="color: #1a73e8; text-decoration: none;">▶️ 回答用フォームを開く</a></div>
          <div style="text-align: center; margin-top: 20px;"><button onclick="google.script.host.close()" style="padding: 8px 24px; background-color: #1a73e8; color: white; border: none; border-radius: 4px; cursor: pointer; font-weight: bold;">閉じる</button></div>
        </div>
      `;
      ui.showModalDialog(HtmlService.createHtmlOutput(htmlContent).setWidth(450).setHeight(530), '✨ 完了');
    } catch (e) {
      ui.alert('⚠️ エラーが発生しました\n\n' + e.message);
    }
  },

  resetHistoryCount: function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    
    const response = ui.alert('履歴のリセット', '現在の「第〇回」という履歴シートの前に「[済]」を付けます。\n（これにより、次回フォーム作成時は再び「第1回」からリスタートします）\n\n実行してよろしいですか？', ui.ButtonSet.YES_NO);
    if (response !== ui.Button.YES) return;

    const allSheets = ss.getSheets();
    let changedCount = 0;
    const regex = /^第(\d+)回/; 

    for (let i = 0; i < allSheets.length; i++) {
      const sheet = allSheets[i];
      const name = sheet.getName();
      if (regex.test(name)) {
        sheet.setName('[済] ' + name);
        changedCount++;
      }
    }

    if (changedCount > 0) {
      ui.alert('✨ リセット完了', `${changedCount}個の履歴シートを過去ログに変更しました。\n次回フォームを作成する際は「第1回」からスタートします！`, ui.ButtonSet.OK);
    } else {
      ui.alert('お知らせ', 'リセット対象となる「第〇回」のシートは見つかりませんでした。\n次回はそのまま「第1回」からスタートします。', ui.ButtonSet.OK);
    }
  },

  setFormDeadline: function() { 
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    
    const confirmResponse = ui.alert(
      '⏳ 締め切り設定の確認', 
      `【 ${Config.DEADLINE_CELL} 】セルに締め切り日を設定しましたか？\n\n※「はい」を押すと、指定した日の 22:00 にフォームが自動で締め切られるよう予約されます。`, 
      ui.ButtonSet.YES_NO
    );
    if (confirmResponse !== ui.Button.YES) return;

    const sheet = ss.getSheetByName(Config.MAIN_SHEET);
    if (!sheet) return;
    const cell = sheet.getRange(Config.DEADLINE_CELL);
    const rawValue = cell.getValue();
    
    if (!rawValue) { 
      ui.alert('⚠️ 入力エラー', `【 ${Config.DEADLINE_CELL} 】セルに締め切りの「日付」が入力されていません。一度キャンセルして日付を設定してください。`, ui.ButtonSet.OK); 
      return; 
    }
    
    let deadlineDate = Object.prototype.toString.call(rawValue) === '[object Date]' ? new Date(rawValue.getTime()) : new Date(String(rawValue));
    
    if (isNaN(deadlineDate.getTime())) { 
      ui.alert('⚠️ 日付エラー', `日付が正しくありません。セルをダブルクリックしてカレンダーから選択してください。`, ui.ButtonSet.OK); 
      return; 
    }
    
    deadlineDate.setHours(22);
    deadlineDate.setMinutes(0);
    deadlineDate.setSeconds(0);
    deadlineDate.setMilliseconds(0);

    if (deadlineDate <= new Date()) { 
      ui.alert('⚠️ 日付エラー', `設定された締め切り日時（${Utilities.formatDate(deadlineDate, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm')}）は過去です。\n未来の日付を選択してください。`, ui.ButtonSet.OK); 
      return; 
    }
    
    let formUrl = null;
    const sheets = ss.getSheets();
    for (let i = 0; i < sheets.length; i++) { if (sheets[i].getName().indexOf('回答データ') !== -1) { formUrl = sheets[i].getFormUrl(); break; } }
    if (!formUrl) { ui.alert('エラー', '連携されているフォームが見つかりません。', ui.ButtonSet.OK); return; }
    
    const triggers = ScriptApp.getProjectTriggers();
    for (let i = 0; i < triggers.length; i++) { if (triggers[i].getHandlerFunction() === 'autoCloseForm') ScriptApp.deleteTrigger(triggers[i]); }
    
    PropertiesService.getScriptProperties().setProperty('TARGET_FORM_URL', formUrl);
    ScriptApp.newTrigger('autoCloseForm').timeBased().at(deadlineDate).create();
    
    cell.setValue(deadlineDate);
    cell.setNumberFormat('"⏳ 締切: "yyyy/MM/dd 22:00');
    cell.clearNote();
    
    ui.alert('✨ 設定完了', Utilities.formatDate(deadlineDate, Session.getScriptTimeZone(), 'yyyy/MM/dd') + ' の【 22:00 】に自動終了をセットしました！', ui.ButtonSet.OK);
  }
};

function autoCloseForm() {
  const formUrl = PropertiesService.getScriptProperties().getProperty('TARGET_FORM_URL');
  if (formUrl) {
    try {
      const form = FormApp.openByUrl(formUrl);
      form.setAcceptingResponses(false);
      form.setCustomClosedFormMessage('提出の締め切り時間を過ぎました。先生の指示に従ってください。');
    } catch (e) {
      console.error('フォームの締め切りに失敗しました。', e);
    }
  }
}