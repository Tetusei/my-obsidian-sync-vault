/**
 * Form.gs
 * Description: フォーム作成・更新、シート自動連携機能（テンプレート色分け対応、空白無視の堅牢マッチング追加、フォーム名シンプル化）
 */

/**
 * フォーム作成・更新、シート自動連携機能
 */
function setupDailyRecordForm() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const baseSheet = ss.getSheetByName(CONFIG.SHEET_BASE);
  const meiboSheet = ss.getSheetByName(CONFIG.SHEET_MEIBO);
  const ui = SpreadsheetApp.getUi();

  try {
    const meiboUrl = `#gid=${meiboSheet.getSheetId()}`;
    const backLinkRichText = SpreadsheetApp.newRichTextValue()
      .setText('↩️ 名簿に戻る')
      .setLinkUrl(meiboUrl)
      .build();

    let templateSheet = ss.getSheetByName(CONFIG.SHEET_TEMPLATE);
    if (!templateSheet) {
      ss.toast('テンプレートシートを作成しています...', '処理中', 2);
      templateSheet = ss.insertSheet(CONFIG.SHEET_TEMPLATE);
      
      templateSheet.getRange('A1').setValue('記録日');
      templateSheet.getRange('B1').setValue('本日の気づき');
      templateSheet.getRange('C1').setRichTextValue(backLinkRichText);
      
      templateSheet.getRange('A1:C1000').applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREEN, true, false);
      
      templateSheet.getRange('A1:C1').setFontWeight('bold');
      templateSheet.setColumnWidth(1, 120);
      templateSheet.setColumnWidth(2, 600);
      templateSheet.setColumnWidth(3, 100);
      templateSheet.setFrozenRows(1);
      
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['〇'], true).setAllowInvalid(true).build();
      templateSheet.getRange('C2:C1000').setDataValidation(rule);
      
      templateSheet.hideSheet();
    } else {
      templateSheet.getRange('C1').setRichTextValue(backLinkRichText);
    }

    const data = meiboSheet.getDataRange().getValues();
    let students = [];
    
    for (let i = 1; i < data.length; i++) {
      const num = data[i][0];
      const name = data[i][1];
      if (num && name) {
        const studentSheetName = `${num}_${name}`;
        students.push(studentSheetName);
        
        const stSheet = ss.getSheetByName(studentSheetName);
        if (stSheet) {
          const cell = meiboSheet.getRange(i + 1, 2);
          const link = SpreadsheetApp.newRichTextValue()
            .setText(name)
            .setLinkUrl(`#gid=${stSheet.getSheetId()}`)
            .build();
          cell.setRichTextValue(link);
          stSheet.getRange('C1').setRichTextValue(backLinkRichText);
        }
      }
    }

    // --- 現在のスプレッドシートとフォームの実際のリンク状態を確認 ---
    const currentLinkedFormUrl = ss.getFormUrl();
    if (!currentLinkedFormUrl) {
      baseSheet.getRange(CONFIG.CELL_FORM_URL).clearContent();
    }
    // -----------------------------------------------------------

    let formUrl = baseSheet.getRange(CONFIG.CELL_FORM_URL).getValue();
    let form = null;
    
    if (formUrl) {
      try {
        const tempForm = FormApp.openByUrl(formUrl);
        const formFile = DriveApp.getFileById(tempForm.getId());
        if (formFile.isTrashed()) {
          form = null;
        } else {
          form = tempForm;
        }
      } catch (e) {
        form = null;
      }
    }

    if (!form) {
      ss.toast('新しいフォームを作成しています...', '処理中', 3);
      
      const now = new Date();
      // 🌟【修正】時刻と「作成:」を消し、日付のみのシンプルなフォーマットに変更
      const timeString = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy/MM/dd');
      form = FormApp.create(`日々の記録フォーム [${timeString}]`);
      
      form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
      
      // 編集用URLを書き込む
      baseSheet.getRange(CONFIG.CELL_FORM_URL).setValue(form.getEditUrl());

      form.addListItem().setTitle('対象生徒').setRequired(true);
      form.addDateItem().setTitle('記録日').setHelpText('※未入力の場合は「送信日」として扱われます。');
      form.addParagraphTextItem().setTitle('本日の気づき').setRequired(true);

      const triggers = ScriptApp.getProjectTriggers();
      const hasTrigger = triggers.some(t => t.getHandlerFunction() === 'onDailyRecordSubmit');
      if (!hasTrigger) {
        ScriptApp.newTrigger('onDailyRecordSubmit')
          .forSpreadsheet(ss)
          .onFormSubmit()
          .create();
      }

      const ssFile = DriveApp.getFileById(ss.getId());
      const formFile = DriveApp.getFileById(form.getId());
      const parents = ssFile.getParents();
      if (parents.hasNext()) {
        parents.next().addFile(formFile);
        DriveApp.getRootFolder().removeFile(formFile); 
      }

      SpreadsheetApp.flush(); 
      Utilities.sleep(2000); 
      
      const allSheets = ss.getSheets();
      for (let i = 0; i < allSheets.length; i++) {
        const fUrl = allSheets[i].getFormUrl();
        if (fUrl && fUrl.indexOf(form.getId()) !== -1) {
          allSheets[i].hideSheet();
          break;
        }
      }
    }

    const items = form.getItems();
    for (let i = 0; i < items.length; i++) {
      if (items[i].getTitle() === '対象生徒') {
        items[i].asListItem().setChoiceValues(students);
        break;
      }
    }

    // 回答用フォームのURLを B16セル に書き込む
    baseSheet.getRange('B16').setValue(form.getPublishedUrl());

    const htmlOutput = HtmlService.createHtmlOutput(`
      <div style="font-family: sans-serif; padding: 10px; color: #333;">
        <p style="color: #1e8e3e; font-weight: bold;">✨ セットアップが完了しました</p>
        <p>「フォームの回答」シートは自動的に非表示になっています。</p>
        <div style="text-align: center; margin-top: 25px; margin-bottom: 20px;">
          <a href="${form.getPublishedUrl()}" target="_blank" style="display: inline-block; padding: 12px 24px; background-color: #1a73e8; color: white; text-decoration: none; border-radius: 8px; font-weight: bold; font-size: 16px; box-shadow: 0 2px 4px rgba(0,0,0,0.2);">
            📝 回答用フォームを開く
          </a>
        </div>
      </div>
    `).setWidth(450).setHeight(250);
    
    ui.showModalDialog(htmlOutput, 'システム連携完了');

  } catch (error) {
    ui.alert('エラーが発生しました:\n' + error.message);
  }
}

function onDailyRecordSubmit(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const namedValues = e.namedValues;
    const student = namedValues['対象生徒'] ? namedValues['対象生徒'][0] : null;
    let recordDate = namedValues['記録日'] ? namedValues['記録日'][0] : null;
    const note = namedValues['本日の気づき'] ? namedValues['本日の気づき'][0] : null;
    const timestamp = namedValues['タイムスタンプ'] ? namedValues['タイムスタンプ'][0] : null;

    if (!student || !note) return;
    if (!recordDate) {
      recordDate = timestamp ? timestamp.split(' ')[0] : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd');
    }

    // 空白（全角・半角）を完全に無視してシートを探す堅牢な検索ロジック
    const normalize = (str) => String(str).replace(/[\s　]/g, '');
    const normalizedStudent = normalize(student);
    
    let sheet = ss.getSheetByName(student); // まずは完全一致で探す
    
    if (!sheet) {
      // 完全一致で見つからない場合、空白を無視して全シートから本人のシートを探す
      const allSheets = ss.getSheets();
      for (let i = 0; i < allSheets.length; i++) {
        if (normalize(allSheets[i].getName()) === normalizedStudent) {
          sheet = allSheets[i];
          break;
        }
      }
    }

    if (!sheet) {
      // それでも見つからない場合（本当にシートが無い場合）のみ新規作成
      const template = ss.getSheetByName(CONFIG.SHEET_TEMPLATE);
      sheet = template.copyTo(ss);
      sheet.setName(student);
      
      // 名簿から性別を探して色分けするロジック（空白無視に対応）
      const meiboSheet = ss.getSheetByName(CONFIG.SHEET_MEIBO);
      const data = meiboSheet.getDataRange().getValues();
      let gender = '';
      for (let i = 1; i < data.length; i++) {
        const meiboName = `${data[i][0]}_${data[i][1]}`;
        if (normalize(meiboName) === normalizedStudent) {
          gender = data[i][2]; // 性別を取得
          break;
        }
      }

      const range = sheet.getRange('A1:C1000');
      range.getBandings().forEach(b => b.remove()); // テンプレートの色をリセット

      let theme = SpreadsheetApp.BandingTheme.LIGHT_GREEN; // デフォルト
      if (gender === '男') {
        theme = SpreadsheetApp.BandingTheme.CYAN;
      } else if (gender === '女') {
        theme = SpreadsheetApp.BandingTheme.PINK;
      }
      range.applyRowBanding(theme, true, false);

      sheet.showSheet(); 
      updateLinkInMeibo(ss, student, sheet);
    }
    
    // 見つかった（または作成された）シートに記録を追記
    sheet.appendRow([recordDate, note]);
  } catch (err) {
    console.error('onDailyRecordSubmit Error: ' + err.message);
  }
}

function updateLinkInMeibo(ss, studentSheetName, sheetObj) {
  const meiboSheet = ss.getSheetByName(CONFIG.SHEET_MEIBO);
  const data = meiboSheet.getDataRange().getValues();
  const sheetUrl = `#gid=${sheetObj.getSheetId()}`;
  
  // 名簿リンク更新時も空白を無視して確実に対象行を探す
  const normalize = (str) => String(str).replace(/[\s　]/g, '');
  const normalizedTarget = normalize(studentSheetName);

  for (let i = 1; i < data.length; i++) {
    if (normalize(`${data[i][0]}_${data[i][1]}`) === normalizedTarget) {
      const cell = meiboSheet.getRange(i + 1, 2);
      // 名簿上の元々の名前（data[i][1]）はそのまま維持してリンクだけ貼る
      const richText = SpreadsheetApp.newRichTextValue().setText(data[i][1]).setLinkUrl(sheetUrl).build();
      cell.setRichTextValue(richText);
      break;
    }
  }
}