/**
 * Admin.gs
 * 管理機能（メニュー追加・シート作成・ダッシュボード更新・メインテナンス）
 * Version: v2.5.0 (全行メニューの追加・onEditによるAPIキー自動退避を追加)
 */

/**
 * スプレッドシートを開いたときにカスタムメニューを追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  const menu1 = ui.createMenu('【1】準備・セットアップ')
    .addItem('📝 個人シートを一括作成（名簿から）', 'createStudentSheets')
    .addItem('📋 日々の記録フォームをセットアップ', 'setupDailyRecordForm');

  const menu2 = ui.createMenu('【2】AIによる所見作成・推敲')
    .addItem('🤖 AI所見を作成（選択行）', 'generateAIComments')
    .addItem('🤖 AI所見を作成（全行）', 'generateAICommentsAll')
    .addSeparator()
    .addItem('🤖 日々の記録から自動生成（選択行）', 'generateFromDailyRecords')
    .addItem('🤖 日々の記録から自動生成（全行）', 'generateFromDailyRecordsAll')
    .addSeparator()
    .addItem('✨ 最終所見を推敲する（選択行）', 'proofreadComments')
    .addItem('✨ 最終所見を推敲する（全行）', 'proofreadCommentsAll');

  const menu3 = ui.createMenu('【3】集計・ダッシュボード')
    .addItem('📊 各学期の記録数（メモ）を名簿に集計する', 'updateRecordCounts');

  const menu4 = ui.createMenu('【4】メインテナンス')
    .addItem('💾 バックアップを作成', 'createBackup')
    .addSeparator()
    .addItem('🗑️ 個人シートを一括削除', 'deleteAllStudentSheets')
    .addItem('🧹 名簿シートの記録をクリア', 'clearMeiboData')
    .addSeparator()
    .addItem('🌸 年度更新を実行（全データ初期化）', 'executeAnnualRenewal');

  ui.createMenu('🤖 AI所見アシスト')
    .addSubMenu(menu1)
    .addSubMenu(menu2)
    .addSubMenu(menu3)
    .addSubMenu(menu4)
    .addToUi();
}

/**
 * 名簿から個人の記録用シートを一括作成
 */
function createStudentSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const meiboSheet = ss.getSheetByName(CONFIG.SHEET_MEIBO);
  const templateSheet = ss.getSheetByName(CONFIG.SHEET_TEMPLATE);
  const ui = SpreadsheetApp.getUi();

  if (!meiboSheet || !templateSheet) {
    return ui.alert('エラー', '名簿シートまたはテンプレートシートが見つかりません。', ui.ButtonSet.OK);
  }

  const data = meiboSheet.getDataRange().getValues();
  let createdCount = 0;
  const meiboUrl = `#gid=${meiboSheet.getSheetId()}`;
  const backLinkRichText = SpreadsheetApp.newRichTextValue().setText('↩️ 名簿に戻る').setLinkUrl(meiboUrl).build();

  ss.toast('個人シートを作成中...', '処理中', 5);

  for (let i = 1; i < data.length; i++) {
    const num = data[i][0];
    const name = data[i][1];
    const gender = data[i][2]; 

    if (!num || !name) continue;

    const sheetName = `${num}_${name}`;
    let targetSheet = ss.getSheetByName(sheetName);
    
    if (!targetSheet) {
      targetSheet = templateSheet.copyTo(ss).setName(sheetName);
      targetSheet.showSheet();
      targetSheet.getRange('A1').setValue(`${num} ${name}`);
      createdCount++;
    }

    const sheetUrl = `#gid=${targetSheet.getSheetId()}`;
    const nameCell = meiboSheet.getRange(i + 1, 2);
    const nameLinkRichText = SpreadsheetApp.newRichTextValue().setText(name).setLinkUrl(sheetUrl).build();
    nameCell.setRichTextValue(nameLinkRichText);
    targetSheet.getRange('C1').setRichTextValue(backLinkRichText);

    const fullRange = targetSheet.getRange('A1:C1000');
    fullRange.getBandings().forEach(b => b.remove());
    let theme = SpreadsheetApp.BandingTheme.LIGHT_GREEN;
    let headerColor = '#b7e1cd';
    if (gender === '男') { theme = SpreadsheetApp.BandingTheme.CYAN; headerColor = '#b3e5fc'; }
    else if (gender === '女') { theme = SpreadsheetApp.BandingTheme.PINK; headerColor = '#f8bbd0'; }
    fullRange.applyRowBanding(theme, true, false);
    targetSheet.getRange('A1:C1').setBackground(headerColor).setFontColor('#000000').setFontWeight('bold');
  }
  ui.alert('完了しました。');
}

/**
 * 記録件数を集計
 */
function updateRecordCounts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const meiboSheet = ss.getSheetByName(CONFIG.SHEET_MEIBO);
  const baseSheet = ss.getSheetByName(CONFIG.SHEET_BASE);
  if (!meiboSheet) return;

  ss.toast('集計中...', '処理中', 3);
  const titles = [['委員会','クラブ・部活','特記事項','1学期件数','2学期件数','3学期件数']];
  meiboSheet.getRange('D1:I1').setValues(titles);

  const term1Start = new Date(baseSheet.getRange(CONFIG.CELL_TERM1_START).getValue());
  const term1End = new Date(baseSheet.getRange(CONFIG.CELL_TERM1_END).getValue());
  const term2Start = new Date(baseSheet.getRange(CONFIG.CELL_TERM2_START).getValue());
  const term2End = new Date(baseSheet.getRange(CONFIG.CELL_TERM2_END).getValue());
  const term3Start = new Date(baseSheet.getRange(CONFIG.CELL_TERM3_START).getValue());
  const term3End = new Date(baseSheet.getRange(CONFIG.CELL_TERM3_END).getValue());

  const data = meiboSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const num = data[i][0]; const name = data[i][1];
    if (!num || !name) continue;
    const stSheet = ss.getSheetByName(`${num}_${name}`);
    let t1=0, t2=0, t3=0;
    if (stSheet) {
      const records = stSheet.getDataRange().getValues();
      for (let r = 1; r < records.length; r++) {
        const d = new Date(records[r][0]);
        if (isNaN(d.getTime())) continue;
        if (d >= term1Start && d <= term1End) t1++;
        else if (d >= term2Start && d <= term2End) t2++;
        else if (d >= term3Start && d <= term3End) t3++;
      }
    }
    meiboSheet.getRange(i + 1, 7, 1, 3).setValues([[t1, t2, t3]]);
  }
  ss.toast('完了しました！');
}

/**
 * バックアップ作成
 */
function createBackup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  ss.toast('バックアップ作成中...', '処理中');
  try {
    const oldFormUrl = ss.getFormUrl();
    let oldFormId = "";
    let oldFormName = "";
    if (oldFormUrl) {
      try { 
        const f = FormApp.openByUrl(oldFormUrl);
        oldFormId = f.getId();
        oldFormName = f.getTitle();
      } catch(e){}
    }

    const file = DriveApp.getFileById(ss.getId());
    const folder = getBackupFolder(file);
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmm');
    const copy = file.makeCopy(`${now}_バックアップ_${ss.getName()}`, folder);
    
    SpreadsheetApp.flush(); 
    Utilities.sleep(3000); 
    
    // 通常のリンク解除トライ
    let cFormUrl = SpreadsheetApp.openById(copy.getId()).getFormUrl();
    for(let i=0; i<3; i++) {
      if (cFormUrl && cFormUrl !== oldFormUrl) break;
      Utilities.sleep(2000);
      cFormUrl = SpreadsheetApp.openById(copy.getId()).getFormUrl();
    }
    if (cFormUrl && cFormUrl !== oldFormUrl) {
      try {
        const f = FormApp.openByUrl(cFormUrl);
        f.removeDestination();
        DriveApp.getFileById(f.getId()).setTrashed(true);
      } catch (e) {} 
    }

    // 🌟 最強の自動お掃除関数を発動（増殖したコピーを一掃）
    sweepGarbageForms(oldFormId, oldFormName || "日々の記録フォーム");

    ui.alert('バックアップを保存箱に作成しました。');
  } catch (e) { ui.alert('エラー: ' + e.message); }
}

/**
 * 個人シート削除
 */
function deleteAllStudentSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('個人シートをすべて削除しますか？', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  const sheets = ss.getSheets();
  const regex = /^\d+_.+/; 
  sheets.forEach(s => { if (regex.test(s.getName())) ss.deleteSheet(s); });

  const meibo = ss.getSheetByName(CONFIG.SHEET_MEIBO);
  if (meibo) {
    const range = meibo.getRange(2, 2, meibo.getLastRow(), 1);
    range.setValues(range.getValues()).setFontColor('#000000').setFontLine('none');
  }
  ui.alert('削除が完了しました。');
}

/**
 * 名簿クリア（1行目の見出し・プルダウン・条件付き書式を完全保護）
 */
function clearMeiboData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const meibo = ss.getSheetByName(CONFIG.SHEET_MEIBO);
  const ui = SpreadsheetApp.getUi();
  
  if (!meibo) return;
  if (ui.alert('確認', '名簿のデータをクリアしますか？\n（1行目の見出し、A列の番号、C列のプルダウン、条件付き書式はそのまま残ります）', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  
  const lastRow = meibo.getLastRow();
  if (lastRow > 1) {
    meibo.getRange(2, 2, lastRow - 1, 2).clearContent();
    meibo.getRange(2, 2, lastRow - 1, 1).setFontColor('#000000').setFontLine('none');
    const maxCols = meibo.getMaxColumns();
    if (maxCols >= 4) {
      meibo.getRange(2, 4, lastRow - 1, maxCols - 3).clearContent();
    }
  }
  
  ss.toast('名簿のクリアが完了しました。', '完了');
}

/**
 * 【完成版・修正】年度更新（徹底お掃除ロジック追加）
 */
function executeAnnualRenewal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    '⚠️ 警告：【年度更新】全データの初期化',
    '【重要】来年度に向けて、現在のデータをすべてリセットします。\n\n' +
    '① 現在の状態を自動でバックアップ保存します。\n' +
    '② 連携中のフォーム本体を削除し、リンクを解除します。\n' +
    '③ 全生徒の個人シートを削除します。\n' +
    '④ 1〜3学期、要録の入力データをすべて消去します。\n' +
    '⑤ 名簿の氏名等(※番号は残します)と、基礎データのURL設定をクリアします。\n\n' +
    '本当に初期化してよろしいですか？（この操作は取り消せません）',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  ss.toast('年度更新用のバックアップを作成中...', '処理中', 10);

  // --- 2. 【年度更新専用のバックアップ処理】 ---
  const oldFormUrl = ss.getFormUrl();
  let oldFormId = "";
  let oldFormName = "";
  if (oldFormUrl) {
    try { 
      const f = FormApp.openByUrl(oldFormUrl);
      oldFormId = f.getId();
      oldFormName = f.getTitle();
    } catch(e){}
  }

  let backupFileName = "";
  try {
    const file = DriveApp.getFileById(ss.getId());
    const folder = getBackupFolder(file);
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmm');
    
    backupFileName = `${now}_[年度更新前データ] ${ss.getName()}`;
    
    const copy = file.makeCopy(backupFileName, folder);
    SpreadsheetApp.flush(); 
    Utilities.sleep(3000); 
    
    let cFormUrl = SpreadsheetApp.openById(copy.getId()).getFormUrl();
    for(let i=0; i<3; i++) {
      if (cFormUrl && cFormUrl !== oldFormUrl) break;
      Utilities.sleep(2000);
      cFormUrl = SpreadsheetApp.openById(copy.getId()).getFormUrl();
    }
    
    if (cFormUrl && cFormUrl !== oldFormUrl) {
      try {
        const cForm = FormApp.openByUrl(cFormUrl);
        cForm.removeDestination();
        DriveApp.getFileById(cForm.getId()).setTrashed(true);
      } catch(e) {}
    }

    // 🌟 最強の自動お掃除関数を発動（増殖したコピーを一掃）
    sweepGarbageForms(oldFormId, oldFormName || "日々の記録フォーム");

  } catch (e) {
    ui.alert('バックアップ作成に失敗したため、処理を中断しました。\n' + e.message);
    return;
  }

  ss.toast('データを初期化しています...', '処理中', 20);

  // 3. フォーム本体（ドライブ上）をゴミ箱へ（元シートのフォーム）
  if (oldFormUrl) {
    try {
      const form = FormApp.openByUrl(oldFormUrl);
      form.removeDestination(); 
      DriveApp.getFileById(form.getId()).setTrashed(true);
    } catch(e) {}
  }

  // 4. 不要シート（回答シート・個人シート）の削除
  const formRegex = /フォームの回答/;
  const studentRegex = /^\d+_.+/;
  SpreadsheetApp.flush();
  
  ss.getSheets().forEach(s => {
    const sName = s.getName();
    if (formRegex.test(sName) || studentRegex.test(sName) || s.getFormUrl()) {
      try { ss.deleteSheet(s); } catch(e) {}
    }
  });

  // 5. 各学期シートのデータクリア
  ['1学期', '2学期', '3学期', '要録'].forEach(name => {
    const s = ss.getSheetByName(name);
    if (!s || s.getMaxRows() <= 1) return;

    const maxRows = s.getMaxRows() - 1;
    const maxCols = s.getMaxColumns();

    if (maxCols > 0) {
      s.getRange(2, 1, maxRows, maxCols).setBackground(null);
    }

    if (name === '要録') {
      if (maxCols >= 7) s.getRange(2, 7, maxRows, 1).clearContent();
      if (maxCols >= 9) s.getRange(2, 9, maxRows, 1).clearContent();
      if (maxCols >= 11) s.getRange(2, 11, maxRows, 1).clearContent(); 
      if (maxCols >= 13) s.getRange(2, 13, maxRows, 1).clearContent(); 
      if (maxCols >= 15) s.getRange(2, 15, maxRows, maxCols - 14).clearContent();
    } else {
      if (maxCols >= 4) s.getRange(2, 4, maxRows, Math.min(6, maxCols - 3)).clearContent();
      if (maxCols >= 11) s.getRange(2, 11, maxRows, 1).clearContent();
      if (maxCols >= 13) s.getRange(2, 13, maxRows, 1).clearContent();
      if (maxCols >= 15) s.getRange(2, 15, maxRows, maxCols - 14).clearContent();
    }
  });

  // 6. 名簿のリセット
  const meibo = ss.getSheetByName(CONFIG.SHEET_MEIBO);
  if (meibo) {
    const lastRow = meibo.getLastRow();
    if (lastRow > 1) {
      meibo.getRange(2, 2, lastRow - 1, 2).clearContent();
      meibo.getRange(2, 2, lastRow - 1, 1).setFontColor('#000000').setFontLine('none');
      const maxCols = meibo.getMaxColumns();
      if (maxCols >= 4) {
        meibo.getRange(2, 4, lastRow - 1, maxCols - 3).clearContent();
      }
    }
  }

  // 7. 基礎データのリセット
  const base = ss.getSheetByName(CONFIG.SHEET_BASE);
  if (base) {
    base.getRange('B15:B16').clearContent();
  }

  ui.alert('🌸 年度更新が完了しました！\n\nバックアップ：' + backupFileName);
}

// 補助関数：バックアップフォルダ取得
function getBackupFolder(file) {
  const parents = file.getParents();
  const root = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  const folderName = '📦_バックアップ保存箱';
  const folders = root.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : root.createFolder(folderName);
}

/**
 * 🌟 補助関数：最強の自動お掃除関数（増殖したコピーフォームを一掃する）
 */
function sweepGarbageForms(activeFormId, baseName) {
  try {
    // 1. 📦_バックアップ保存箱 の中のフォームは、問答無用で全消去（ここはシート専用の箱なので）
    const root = DriveApp.getRootFolder();
    const backupFolders = root.getFoldersByName('📦_バックアップ保存箱');
    if (backupFolders.hasNext()) {
      const bFolder = backupFolders.next();
      const forms = bFolder.getFilesByType(MimeType.GOOGLE_FORMS);
      while(forms.hasNext()) {
        const f = forms.next();
        if (f.getId() !== activeFormId) f.setTrashed(true);
      }
    }

    // 2. 現在のスプレッドシートと同じフォルダ内にある「増殖したフォーム」を消去
    const ssFile = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
    const parents = ssFile.getParents();
    if (parents.hasNext()) {
      const pFolder = parents.next();
      const forms = pFolder.getFilesByType(MimeType.GOOGLE_FORMS);
      while(forms.hasNext()) {
        const f = forms.next();
        if (f.getId() !== activeFormId) {
          const name = f.getName();
          // 本物のフォーム名（例：「日々の記録フォーム」）が名前に含まれていればすべてゴミ箱へ
          if (baseName && name.includes(baseName)) {
            f.setTrashed(true);
          }
        }
      }
    }
  } catch (e) {
    console.error("Garbage collection failed: " + e.message);
  }
}

/**
 * セルが編集されたときに実行されるイベントハンドラ
 */
function onEdit(e) {
  if (!e) return;
  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();

  // 基礎データシートの編集のみ対象とする
  if (sheetName !== CONFIG.SHEET_BASE) return;

  // 単一セルの編集のみ対象とする
  if (range.getNumRows() !== 1 || range.getNumColumns() !== 1) return;

  const row = range.getRow();
  const col = range.getColumn();

  // B1, B2, B3 セルの判定
  // B1 = (1, 2), B2 = (2, 2), B3 = (3, 2)
  if (col === 2 && row >= 1 && row <= 3) {
    const value = range.getValue().toString().trim();
    const propKey = `GEMINI_API_KEY_${row}`;
    const scriptProperties = PropertiesService.getScriptProperties();

    if (value === "") {
      // セルが空にされた場合は、対応するスクリプトプロパティのAPIキーを削除
      scriptProperties.deleteProperty(propKey);
    } else if (value !== "設定済み") {
      // 設定済み以外の値（生のAPIキー）が入力された場合、スクリプトプロパティに格納
      scriptProperties.setProperty(propKey, value);
      // セルの表示を「設定済み」に変更
      range.setValue("設定済み");
    }
  }
}