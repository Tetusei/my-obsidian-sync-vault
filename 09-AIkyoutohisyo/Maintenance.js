/**
 * 過去の重複ToDoを一掃するスクリプト（使い切り）
 * 同じ日付・同じ件名のタスクが2つある場合、hachu7010 からのメール（または重複の片方）を削除します
 */
function cleanUpPastDuplicateTasks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const todoSheet = ss.getSheetByName(Config.SHEET_NAME_TODO);
  if (!todoSheet) {
    SpreadsheetApp.getUi().alert("エラー: ToDoシートが見つかりません。");
    return;
  }

  // ToDoシートの全データを取得
  const lastRow = todoSheet.getLastRow();
  if (lastRow < 2) {
    ss.toast("データがありませんでした。", "🧹 お掃除完了", 3000);
    return;
  }
  
  const data = todoSheet.getDataRange().getValues();
  let deleteCount = 0;
  
  // 日付の表記揺れを "yyyy/MM/dd" に統一して比較するヘルパー
  const normalizeDate = (val) => {
    if (val instanceof Date) {
      return Utilities.formatDate(val, "JST", "yyyy/MM/dd");
    }
    if (val && typeof val === 'string') {
      const match = val.trim().match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
      if (match) {
        const y = match[1];
        const m = match[2].padStart(2, '0');
        const d = match[3].padStart(2, '0');
        return `${y}/${m}/${d}`;
      }
    }
    return String(val || "").trim();
  };

  // 行の削除を行うため、必ず「下（末尾）から逆順に」ループを回します
  for (let i = lastRow - 1; i >= 1; i--) {
    const rowDate = data[i][Config.TODO_COL.DATE];
    const rowTitle = data[i][Config.TODO_COL.TITLE] ? data[i][Config.TODO_COL.TITLE].toString().trim() : "";
    const rowContent = data[i][Config.TODO_COL.CONTENT] ? data[i][Config.TODO_COL.CONTENT].toString() : "";
    const rowMemo = data[i][Config.TODO_COL.MEMO] ? data[i][Config.TODO_COL.MEMO].toString() : "";
    
    if (!rowTitle) continue;

    // 現在の行と同じ「日付」かつ「件名」を持つ行が、他に存在するかチェック
    let isDuplicate = false;
    let matchIndex = -1;
    
    for (let j = 1; j < data.length; j++) {
      if (i === j) continue;
      
      const compareDate = data[j][Config.TODO_COL.DATE];
      const compareTitle = data[j][Config.TODO_COL.TITLE] ? data[j][Config.TODO_COL.TITLE].toString().trim() : "";
      
      if (rowTitle === compareTitle && normalizeDate(rowDate) === normalizeDate(compareDate)) {
        isDuplicate = true;
        matchIndex = j;
        break;
      }
    }

    if (isDuplicate) {
      const rowHas7010 = rowContent.includes("hachu7010") || rowMemo.includes("hachu7010");
      const compareContent = data[matchIndex][Config.TODO_COL.CONTENT] ? data[matchIndex][Config.TODO_COL.CONTENT].toString() : "";
      const compareMemo = data[matchIndex][Config.TODO_COL.MEMO] ? data[matchIndex][Config.TODO_COL.MEMO].toString() : "";
      const compareHas7010 = compareContent.includes("hachu7010") || compareMemo.includes("hachu7010");

      if (rowHas7010) {
        todoSheet.deleteRow(i + 1);
        data.splice(i, 1);
        deleteCount++;
      } 
      else if (compareHas7010) {
        continue;
      } 
      else if (i > matchIndex) {
        todoSheet.deleteRow(i + 1);
        data.splice(i, 1);
        deleteCount++;
      }
    }
  }

  ss.toast(`過去の重複ToDoを ${deleteCount} 件削除し、綺麗に整理しました！`, "🧹 お掃除完了", 4000);
}

/**
 * 🧹 重複ToDoを自動検出してモーダルダイアログで一覧表示する（B案）
 */
function showDuplicateTasksDialog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(Config.SHEET_NAME_TODO);
  if (!sheet) {
    SpreadsheetApp.getUi().alert("エラー: ToDoシートが見つかりません。");
    return;
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("ToDoデータが存在しません。");
    return;
  }
  
  const data = sheet.getDataRange().getValues();
  const groups = {}; // { normalizedSubject: [ {rowIdx, date, title, pic, status, memo, mailLink}, ... ] }
  
  // データ解析
  for (let i = 1; i < data.length; i++) {
    const rowIdx = i + 1; // 1-based行番号
    const titleFormula = String(data[i][Config.TODO_COL.TITLE] || "");
    const dateVal = data[i][Config.TODO_COL.DATE];
    const pic = data[i][Config.TODO_COL.PIC] || "";
    const status = data[i][Config.TODO_COL.STATUS] || "";
    const memo = String(data[i][Config.TODO_COL.MEMO] || "").substring(0, 100);
    const mailLink = String(data[i][Config.TODO_COL.MAIL_LINK] || "").trim();
    
    if (!titleFormula) continue;
    
    // Hyperlink式からタイトルを抽出
    let title = titleFormula;
    const linkMatch = titleFormula.match(/=HYPERLINK\("[^"]*",\s*"([^"]*)"\)/i);
    if (linkMatch) {
      title = linkMatch[1];
    }
    title = title.replace(/^📎\s*/, "");
    
    const normSub = normalizeSubject(title);
    if (!normSub) continue;
    
    const dateStr = dateVal instanceof Date ? Utilities.formatDate(dateVal, "JST", "MM/dd") : String(dateVal || "");
    
    if (!groups[normSub]) {
      groups[normSub] = [];
    }
    groups[normSub].push({
      rowIdx: rowIdx,
      date: dateStr,
      title: title,
      pic: pic,
      status: status,
      memo: memo,
      mailLink: mailLink
    });
  }
  
  // 重複しているもの（2件以上）だけを抽出
  const duplicates = [];
  for (const sub in groups) {
    if (groups[sub].length >= 2) {
      duplicates.push({
        subject: sub,
        items: groups[sub]
      });
    }
  }
  
  // ダイアログHTMLの生成
  let html = `
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body {
          font-family: 'Helvetica Neue', Arial, 'Hiragino Kaku Gothic ProN', Meiryo, sans-serif;
          margin: 0;
          padding: 20px;
          color: #333;
          background-color: #f8f9fa;
        }
        h2 {
          font-size: 18px;
          margin-top: 0;
          color: #1a73e8;
          border-bottom: 2px solid #e8f0fe;
          padding-bottom: 8px;
        }
        .desc {
          font-size: 12px;
          color: #5f6368;
          margin-bottom: 20px;
          line-height: 1.5;
        }
        .dup-group {
          background-color: #fff;
          border: 1px solid #dadce0;
          border-radius: 8px;
          margin-bottom: 16px;
          box-shadow: 0 1px 2px 0 rgba(60,64,67,0.3), 0 1px 3px 1px rgba(60,64,67,0.15);
          overflow: hidden;
        }
        .group-header {
          background-color: #f1f3f4;
          padding: 10px 16px;
          font-size: 13px;
          font-weight: bold;
          color: #202124;
          border-bottom: 1px solid #dadce0;
        }
        .item-row {
          display: flex;
          align-items: center;
          padding: 12px 16px;
          border-bottom: 1px solid #f1f3f4;
          font-size: 13px;
        }
        .item-row:last-child {
          border-bottom: none;
        }
        .item-row:hover {
          background-color: #f8f9fa;
        }
        .checkbox-container {
          margin-right: 12px;
        }
        .item-info {
          flex-grow: 1;
        }
        .item-meta {
          font-size: 11px;
          color: #5f6368;
          margin-top: 4px;
        }
        .badge {
          display: inline-block;
          padding: 2px 6px;
          border-radius: 4px;
          font-size: 10px;
          font-weight: bold;
          margin-right: 6px;
        }
        .badge-status-completed {
          background-color: #e6f4ea;
          color: #137333;
        }
        .badge-status-active {
          background-color: #fef7e0;
          color: #b06000;
        }
        .mail-link {
          color: #1a73e8;
          text-decoration: none;
          font-weight: bold;
        }
        .mail-link:hover {
          text-decoration: underline;
          color: #1557b0;
        }
        .btn-container {
          position: sticky;
          bottom: 0;
          background-color: #f8f9fa;
          padding: 16px 0 0 0;
          border-top: 1px solid #dadce0;
          display: flex;
          justify-content: flex-end;
          gap: 12px;
        }
        button {
          padding: 8px 16px;
          border-radius: 4px;
          font-size: 13px;
          font-weight: bold;
          cursor: pointer;
          border: none;
        }
        .btn-cancel {
          background-color: transparent;
          color: #5f6368;
          border: 1px solid #dadce0;
        }
        .btn-cancel:hover {
          background-color: #f1f3f4;
        }
        .btn-delete {
          background-color: #d93025;
          color: #fff;
        }
        .btn-delete:hover {
          background-color: #b31412;
        }
        .no-dup {
          text-align: center;
          padding: 40px;
          color: #5f6368;
          font-size: 14px;
        }
      </style>
    </head>
    <body>
      <h2>🧹 重複ToDoの整理（クリーンアップ）</h2>
      <p class="desc">
        同じメール（件名が類似しているもの）が複数登録されている候補を表示しています。<br>
        <strong>タイトル（青い文字）をクリックすると、実際のGmailを開いて内容を確認できます。</strong><br>
        削除したい方のタスクにチェックを入れて、下の「選択したToDoを削除」ボタンを押してください。
      </p>
      
      <div id="content">
  `;
  
  if (duplicates.length === 0) {
    html += '<div class="no-dup">🎉 重複している同一メールのToDoは見つかりませんでした。</div>';
  } else {
    duplicates.forEach(group => {
      html += `
        <div class="dup-group">
          <div class="group-header">件名: ${group.subject}</div>
      `;
      group.items.forEach(item => {
        const isCompleted = item.status.includes('完了');
        const badgeClass = isCompleted ? 'badge-status-completed' : 'badge-status-active';
        const badgeLabel = isCompleted ? '完了' : item.status || '未着手';
        
        // メールリンクがあり、かつ正しいURL形式（httpから始まり、改行や添付マークを含まない）の場合のみリンク化
        const isLinkValid = item.mailLink && item.mailLink.toString().startsWith('http') && !item.mailLink.toString().includes('\n') && !item.mailLink.toString().includes('\r') && !item.mailLink.toString().includes('📎');
        
        const titleHtml = isLinkValid 
          ? `<a href="${item.mailLink}" target="_blank" class="mail-link" title="Gmailでメールを開く">行 ${item.rowIdx}: ${item.title} 🔗</a>`
          : `行 ${item.rowIdx}: ${item.title}`;
        
        html += `
          <div class="item-row">
            <div class="checkbox-container">
              <input type="checkbox" class="task-checkbox" data-row="${item.rowIdx}" value="${item.rowIdx}">
            </div>
            <div class="item-info">
              <div>${titleHtml}</div>
              <div class="item-meta">
                <span class="badge ${badgeClass}">${badgeLabel}</span>
                <span>日付: ${item.date}</span> | 
                <span>担当: ${item.pic}</span>
                ${item.memo ? ` | <span style="color:#80868b;">メモ: ${item.memo}</span>` : ''}
              </div>
            </div>
          </div>
        `;
      });
      html += `</div>`;
    });
  }
  
  html += `
      </div>
      
      <div class="btn-container">
        <button class="btn-cancel" onclick="google.script.host.close()">閉じる</button>
        ${duplicates.length > 0 ? `<button class="btn-delete" onclick="deleteSelectedTasks()">選択したToDoを削除</button>` : ''}
      </div>
      
      <script>
        function deleteSelectedTasks() {
          const checkboxes = document.querySelectorAll('.task-checkbox:checked');
          const rowIndexes = [];
          checkboxes.forEach(cb => {
            rowIndexes.push(parseInt(cb.value, 10));
          });
          
          if (rowIndexes.length === 0) {
            alert('削除するToDoにチェックを入れてください。');
            return;
          }
          
          if (!confirm(rowIndexes.length + ' 件のToDoを削除します。よろしいですか？\\\\n※削除前に自動でバックアップが保存されます。')) {
            return;
          }
          
          // ボタンを無効化
          const btn = document.querySelector('.btn-delete');
          btn.disabled = true;
          btn.innerText = '処理中...';
          
          google.script.run
            .withSuccessHandler(function(msg) {
              alert(msg);
              google.script.host.close();
              // リフレッシュのために再表示
              google.script.run.showDuplicateTasksDialog();
            })
            .withFailureHandler(function(err) {
              alert('エラーが発生しました: ' + err);
              btn.disabled = false;
              btn.innerText = '選択したToDoを削除';
            })
            .deleteSelectedTodoRows(rowIndexes);
        }
      </script>
    </body>
    </html>
  `;
  
  const userInterface = HtmlService.createHtmlOutput(html)
    .setWidth(750)
    .setHeight(550);
  
  SpreadsheetApp.getUi().showModalDialog(userInterface, "重複ToDoのクリーンアップ");
}

/**
 * 選択された複数の行番号を安全に削除する（自動バックアップ付き）
 * @param {number[]} rowIndexes 削除対象の行番号配列（1-indexed）
 */
function deleteSelectedTodoRows(rowIndexes) {
  if (!rowIndexes || rowIndexes.length === 0) return "削除対象の指定がありません。";
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. 自動バックアップの作成
  try {
    createBackupCore(ss);
  } catch(bErr) {
    console.error("バックアップ作成エラー: " + bErr);
  }
  
  const sheet = ss.getSheetByName(Config.SHEET_NAME_TODO);
  if (!sheet) throw new Error("ToDoシートが見つかりません。");
  
  // 行インデックスを降順（下から順）にソートします
  const sortedIndexes = rowIndexes.sort((a, b) => b - a);
  
  sortedIndexes.forEach(rowIdx => {
    sheet.deleteRow(rowIdx);
  });
  
  return `選択された ${rowIndexes.length} 件の重複ToDoを正常に削除し、バックアップを保存しました。`;
}