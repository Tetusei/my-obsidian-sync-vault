/**
 * Admin.gs
 * VERSION: v49.4 (重複排除・トリガー管理機能追加)
 * ・メニュー作成、ダイアログ表示、トリガー管理を担当
 */

/**
 * メニュー作成
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🏫 校務支援')
    .addSubMenu(ui.createMenu('📅 行事予定管理')
      .addItem('1. 行事PDFから取込 (AI)', 'showEventPdfDialog')
      .addItem('2. カレンダーへ登録', 'showEventRegDialog'))
    .addSeparator()
    .addSubMenu(ui.createMenu('🍛 給食献立管理')
      .addItem('1. 給食PDFから取込 (AI)', 'showMenuPdfDialog')
      .addItem('2. カレンダーへ登録', 'showMenuRegDialog'))
    .addSeparator()
    .addSubMenu(ui.createMenu('🤖 日次ボット')
      .addItem('🔄 予定表を更新 (今日)', 'forceUpdateToToday') // Main.gsの関数を呼ぶ
      .addItem('📤 今の内容でチャット送信', 'sendManualChat'))    // Main.gsの関数を呼ぶ
    .addSeparator()
    .addSubMenu(ui.createMenu('メインテナンス')  
    .addItem('🌍 地域コード設定', 'updateAreaCode')          // CalendarManager.gsの関数を呼ぶ
      .addItem('📦 システムのバックアップ', 'createSystemBackup'))
    .addSeparator()
    .addSubMenu(ui.createMenu('⚙️ 初期設定')
      .addItem('▶️ 自動化トリガーをONにする', 'enableTriggers')
      .addItem('⏹️ 自動化トリガーをOFFにする', 'disableTriggers')
      .addSeparator()
      .addItem('【ON】11時〜12時の自動トリガーをセットする', 'setupWbgtTrigger')
      .addItem('【OFF】自動トリガーを解除（削除）する', 'clearWbgtTrigger'))
    .addToUi();
}

// ==========================================
//  トリガー管理機能 (ON / OFF)
// ==========================================

/**
 * トリガーを一括ONにする
 */
function enableTriggers() {
  // 重複登録を防ぐため、一度すべてOFFにする（裏側で静かに実行）
  disableTriggers(true);

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    // 1. respondToEdit: スプレッドシート編集時
    ScriptApp.newTrigger('respondToEdit')
      .forSpreadsheet(ss)
      .onEdit()
      .create();

    // 2. sendDailyChatNotification: 毎日午前7時〜8時
    ScriptApp.newTrigger('sendDailyChatNotification')
      .timeBased()
      .everyDays(1)
      .atHour(7)
      .create();

    Browser.msgBox(
      "設定完了", 
      "自動化トリガーを【ON】にしました。\\n\\n✅ シートの日付変更時の自動更新\\n✅ 毎朝7時〜8時のBot自動送信\\n\\nが有効になっています。", 
      Browser.Buttons.OK
    );
  } catch (e) {
    Browser.msgBox("エラー", "トリガーの設定に失敗しました。権限を確認してください。\\n" + e.message, Browser.Buttons.OK);
  }
}

/**
 * トリガーを一括OFFにする
 * @param {boolean} isSilent - 完了メッセージを非表示にするかどうか（システム内部からの呼び出し用）
 */
function disableTriggers(isSilent) {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    let count = 0;
    
    for (let i = 0; i < triggers.length; i++) {
      const handlerName = triggers[i].getHandlerFunction();
      // 今回対象とする2つのトリガーのみを削除する
      if (handlerName === 'respondToEdit' || handlerName === 'sendDailyChatNotification') {
        ScriptApp.deleteTrigger(triggers[i]);
        count++;
      }
    }
    
    if (isSilent !== true) {
      Browser.msgBox(
        "設定解除", 
        `自動化トリガーを【OFF】にしました。\\n（解除した設定: ${count} 件）\\n\\n自動更新および毎朝の自動送信は停止しています。`, 
        Browser.Buttons.OK
      );
    }
  } catch (e) {
    if (isSilent !== true) {
      Browser.msgBox("エラー", "トリガーの解除に失敗しました。\\n" + e.message, Browser.Buttons.OK);
    }
  }
}


// ==========================================
//  以下、ダイアログ表示用関数・バックアップ機能
// ==========================================

function showEventRegDialog() { showRegDialog('Event', '行事'); }
function showMenuRegDialog() { showRegDialog('Menu', '給食'); }
function showEventPdfDialog() { showPdfDialog('Event', '📄 行事予定PDF取込'); }
function showMenuPdfDialog() { showPdfDialog('Menu', '🍛 給食献立PDF取込'); }

/**
 * 分割登録ダイアログ
 */
function showRegDialog(mode, filterTxt) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets().filter(s => s.getName().includes(filterTxt)).map(s => s.getName());
  
  if (sheets.length === 0) {
    Browser.msgBox(`「${filterTxt}」を含むシートが見つかりません。`);
    return;
  }

  const html = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body{font-family:sans-serif;padding:15px;color:#333;font-size:14px;}
          .list{margin:10px 0;max-height:100px;overflow-y:auto;border:1px solid #eee;padding:5px;background:#fff;}
          label{display:block;margin:3px 0;font-size:13px;cursor:pointer;}
          button{width:100%;padding:12px;background:#1a73e8;color:white;border:none;border-radius:4px;cursor:pointer;font-weight:bold;font-size:14px;}
          button:hover{background:#1557b0;}
          button:disabled{background:#ccc;cursor:not-allowed;}
          #status-box{margin-top:15px;padding:12px;border-radius:4px;font-size:13px;border:1px solid #ddd;background:#f9f9f9;}
          .progress-bar{height:8px;background:#eee;border-radius:4px;margin:10px 0;overflow:hidden;}
          .progress-fill{height:100%;background:#1a73e8;width:0%;transition:width 0.3s;}
          .detail{font-size:11px;color:#666;margin-top:5px;}
          .success-msg{color:#188038; font-weight:bold; font-size:1.1em;}
        </style>
      </head>
      <body>
        <strong>登録対象を選択：</strong>
        <div class="list">
          ${sheets.map(s => `<label><input type="checkbox" name="sheetName" value="${s}" checked> ${s}</label>`).join('')}
        </div>
        <button id="btn" onclick="runSplit()">カレンダー登録開始</button>
        <div id="status-box">
          <div id="main-msg">待機中</div>
          <div class="progress-bar"><div id="bar" class="progress-fill"></div></div>
          <div id="detail-msg" class="detail"></div>
        </div>
        <script>
          async function runSplit() {
            const btn = document.getElementById('btn');
            const mainMsg = document.getElementById('main-msg');
            const detailMsg = document.getElementById('detail-msg');
            const bar = document.getElementById('bar');
            const names = Array.from(document.querySelectorAll('input[name="sheetName"]:checked')).map(c => c.value);
            if(names.length === 0) { mainMsg.innerText = "❌ シートを選択してください"; return; }
            btn.disabled = true; btn.innerText = "処理中...";
            let totalSuccess = 0;
            try {
              for (let i = 0; i < names.length; i++) {
                const name = names[i];
                mainMsg.innerHTML = "⏳ <b>" + name + "</b> を処理中 (" + (i+1) + "/" + names.length + ")";
                let startRow = 1; let hasMore = true; let step = 5;
                while(hasMore) {
                  detailMsg.innerText = startRow + "行目付近をスキャン中... (現在の登録数: " + totalSuccess + "件)";
                  const result = await new Promise((resolve, reject) => {
                    google.script.run.withSuccessHandler(resolve).withFailureHandler(reject).batchRegSplit(name, '${mode}', startRow, step); 
                  });
                  if (result.count) totalSuccess += result.count;
                  if (result.isFinished) hasMore = false; else { startRow = result.nextRow; bar.style.width = Math.min(95, (startRow / 40) * 100) + "%"; }
                }
                bar.style.width = ((i+1) / names.length * 100) + "%";
              }
              if (totalSuccess > 0) { mainMsg.innerHTML = "<span class='success-msg'>✅ 登録完了: " + totalSuccess + " 件</span>"; detailMsg.innerText = "カレンダーへの反映を確認してください。"; }
              else { mainMsg.innerHTML = "⚠️ <b>登録対象なし</b>"; detailMsg.innerText = "チェックされた行、または未登録の行がありませんでした。"; }
            } catch(e) { mainMsg.innerText = "❌ エラーが発生しました"; detailMsg.innerText = e.message; }
            finally { btn.innerText = "閉じる"; btn.disabled = false; btn.onclick = () => google.script.host.close(); }
          }
        </script>
      </body>
    </html>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setHeight(380).setWidth(400), `カレンダー登録(${mode})`);
}

/**
 * PDF取込ダイアログ
 */
function showPdfDialog(mode, title) {
  const html = `
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body{font-family:sans-serif;padding:15px;color:#333;background-color:#f8f9fa;}
        h3{font-size:16px;margin-top:0;color:#1a73e8;}
        select,button{width:100%;margin-bottom:15px;padding:12px;border-radius:4px;box-sizing:border-box;}
        select{border:1px solid #ccc;background:#fff;}
        button{background:#1a73e8;color:#fff;border:none;font-weight:bold;cursor:pointer;box-shadow:0 1px 2px rgba(0,0,0,0.1);}
        button:disabled{background:#ccc;cursor:not-allowed;}
        #s{font-size:13px;color:#666;white-space:pre-wrap;line-height:1.4;margin-top:10px;}
      </style>
    </head>
    <body>
      <h3 id="title">PDFを選択してください</h3>
      <select id="f"><option>読込中...</option></select>
      <button id="b" onclick="run()">解析実行</button>
      <div id="s"></div>
      <script>
        window.mode = "${mode}";
        google.script.run.withSuccessHandler(l=>{
          const s=document.getElementById('f'); s.innerHTML='';
          if(l && l.length && l[0].id !== "") l.forEach(x=>s.add(new Option(x.name,x.id)));
          else { s.add(new Option(l[0] ? l[0].name : "PDFなし","")); document.getElementById('b').disabled = true; }
        }).getClientFileList(); 
        function run(){
          const b=document.getElementById('b'); const fileId = document.getElementById('f').value;
          if(!fileId) return;
          b.disabled=true; b.innerText="AI解析中... (30秒ほど)";
          document.getElementById('s').innerText = "Gemini ProがPDFを解析しています...\\nそのままお待ちください。";
          google.script.run.withSuccessHandler(r=>{
            document.getElementById('s').innerText = r ? r.msg : "解析完了"; b.innerText = "完了";
            if(r && !r.error) setTimeout(() => google.script.host.close(), 4000);
          }).processPdf(fileId, window.mode);
        }
      </script>
    </body>
    </html>
  `;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setHeight(350).setWidth(400), title);
}

/**
 * システムバックアップ
 */
function createSystemBackup() {
  const BACKUP_FOLDER_NAME = "📦_バックアップ保存箱";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    const currentFile = DriveApp.getFileById(ss.getId());
    const parents = currentFile.getParents();
    const parentFolder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
    const folders = parentFolder.getFoldersByName(BACKUP_FOLDER_NAME);
    const targetFolder = folders.hasNext() ? folders.next() : parentFolder.createFolder(BACKUP_FOLDER_NAME);
    const dateStr = Utilities.formatDate(new Date(), "JST", "yyyyMMdd_HHmm");
    const backupName = `${ss.getName()}_${dateStr}`;
    currentFile.makeCopy(backupName, targetFolder);
    Browser.msgBox("バックアップ完了", `「${BACKUP_FOLDER_NAME}」に保存しました。\\n\\n📄 ${backupName}`, Browser.Buttons.OK);
  } catch (e) {
    console.error(e);
    Browser.msgBox("エラー", `バックアップに失敗しました。\\n${e.message}`, Browser.Buttons.OK);
  }
}