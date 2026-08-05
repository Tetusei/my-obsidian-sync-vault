/**
 * ファイル名: DataSync.gs
 * 役割: 成績集計・データバックアップ作成
 * バージョン: v5.0.0 (罫線自動描画対応)
 */

const DataSync = {
  createBackup: function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const parentFolder = DriveApp.getFileById(ss.getId()).getParents().next();
    const folders = parentFolder.getFoldersByName(Config.BACKUP_FOLDER_NAME);
    const backupFolder = folders.hasNext() ? folders.next() : parentFolder.createFolder(Config.BACKUP_FOLDER_NAME);
    const backupName = `[${Config.VERSION}]_バックアップ_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss')}_${ss.getName()}`;
    DriveApp.getFileById(ss.getId()).makeCopy(backupName, backupFolder);
  },

  syncScoresToRoster: function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const rosterSheet = ss.getSheetByName(Config.ROSTER_SHEET);
    const ui = SpreadsheetApp.getUi();

    if (!rosterSheet) {
      ui.alert('エラー', '名簿シートが見つかりません。', ui.ButtonSet.OK);
      return;
    }

    const allSheets = ss.getSheets();
    const historySheets = allSheets.filter(s => {
      return /^第(\d+)回/.test(s.getName());
    }).sort((a, b) => {
      const aNum = parseInt(a.getName().match(/第(\d+)回/)[1]);
      const bNum = parseInt(b.getName().match(/第(\d+)回/)[1]);
      return aNum - bNum;
    });

    if (historySheets.length === 0) {
      ui.alert('お知らせ', '集計対象となる履歴シート（第1回...など）が見つかりませんでした。', ui.ButtonSet.OK);
      return;
    }

    ss.toast('スコアを集計中...', '📊 集計中', -1);

    const rosterData = rosterSheet.getDataRange().getValues();
    const studentMap = {}; 
    for (let i = 1; i < rosterData.length; i++) {
      const num = rosterData[i][0];
      if (num) studentMap[num] = i + 1;
    }

    const lastCol = rosterSheet.getLastColumn();
    if (lastCol >= 4) {
      rosterSheet.getRange(1, 4, rosterSheet.getMaxRows(), lastCol - 3).clearContent();
    }

    historySheets.forEach((histSheet, index) => {
      const currentCol = 4 + index; 
      const sheetName = histSheet.getName();
      const sheetId = histSheet.getSheetId();

      const linkFormula = `=HYPERLINK("#gid=${sheetId}", "${sheetName}")`;
      rosterSheet.getRange(1, currentCol).setFormula(linkFormula);
      rosterSheet.getRange(1, currentCol).setBackground('#e1f5fe').setFontWeight('bold');

      const histData = histSheet.getDataRange().getValues();
      
      for (let j = Config.DATA_START_ROW - 1; j < histData.length; j++) {
        const sNum = histData[j][0]; 
        let sScore = histData[j][Config.COL_SCORE - 1]; 
        
        if (sNum && studentMap[sNum]) {
          if (Object.prototype.toString.call(sScore) === '[object Date]') {
            const m = sScore.getMonth() + 1;
            const d = sScore.getDate();      
            sScore = m + '/' + d + '点';     
          } else if (sScore && !sScore.toString().includes('点') && sScore.toString().includes('/')) {
            sScore = sScore + '点';
          }

          const targetCell = rosterSheet.getRange(studentMap[sNum], currentCol);
          targetCell.setNumberFormat('@'); 
          targetCell.setValue(sScore);
        }
      }
    });

    rosterSheet.autoResizeColumns(4, historySheets.length);
    
    // ▼ 自動罫線描画ロジック ▼
    const finalLastRow = rosterSheet.getLastRow();
    const finalLastCol = 3 + historySheets.length; 
    
    rosterSheet.getDataRange().setBorder(false, false, false, false, false, false);
    
    if (finalLastRow > 0 && finalLastCol > 0) {
      rosterSheet.getRange(1, 1, finalLastRow, finalLastCol).setBorder(
        true, true, true, true, true, true, 
        '#000000', 
        SpreadsheetApp.BorderStyle.SOLID
      );
    }

    ss.toast('すべてのスコアの集計と書式設定が完了しました！', '✨ 完了', 5);
  }
};