/**
 * ファイル名: Security.gs
 * 役割: シートの保護・ロック管理
 * バージョン: v5.0.0
 */

const Security = {
  checkProtectionStatus: function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    const targetSheet = ss.getSheetByName(Config.MAIN_SHEET); 
    if (!targetSheet) return;
    if (targetSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).length === 0) {
      ui.alert('🔓 現在の状態：【 ロック解除中 】', '生徒がどのセルでも自由に編集できてしまう状態です。', ui.ButtonSet.OK);
    } else {
      ui.alert('🔒 現在の状態：【 ロック作動中 】', '生徒は「自分の行の指定された場所」以外は編集できない、安全な状態です。', ui.ButtonSet.OK);
    }
  },

  setupRosterLink: function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    const targetSheet = ss.getSheetByName(Config.MAIN_SHEET); 
    if (!targetSheet) return;

    if (targetSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).length > 0) {
      ui.alert('お知らせ', '「添削」シートはすでにロック作動中です。', ui.ButtonSet.OK); return;
    }
    const response = ui.alert('確認', '「添削」シートに保護（ロック）をセットアップし、生徒の編集を制限します。\nよろしいですか？', ui.ButtonSet.YES_NO);
    if (response !== ui.Button.YES) return;
    
    ss.toast('セットアップを開始します...', '🔒 処理中', -1);
    const rosterSheet = ss.getSheetByName(Config.ROSTER_SHEET);
    if (!rosterSheet) return;
    const rosterData = rosterSheet.getDataRange().getValues();
    const students = [];
    for(let i=1; i<rosterData.length; i++) {
      if (rosterData[i][0]) students.push({num: rosterData[i][0], email: rosterData[i][1], name: rosterData[i][2]});
    }
    
    Security.clearProtections_(targetSheet); // ★分割に伴い Security に変更
    const teacherEmail = Session.getEffectiveUser().getEmail();

    for(let i=0; i<students.length; i++) {
      const rowNum = i + Config.DATA_START_ROW; 
      
      targetSheet.getRange(rowNum, Config.COL_NUM).setValue(students[i].num);
      if(students[i].name) targetSheet.getRange(rowNum, Config.COL_STUDENT_NAME).setValue(students[i].name);

      const studentRange = targetSheet.getRange(rowNum, Config.COL_STUDENT); 
      const protection = studentRange.protect().setDescription(`${students[i].num}番`);
      protection.removeEditors(protection.getEditors());
      protection.addEditor(teacherEmail); 
      if (students[i].email) { try { protection.addEditor(students[i].email); } catch (e) {} }
    }

    const lastRow = Math.max(students.length + Config.DATA_START_ROW - 1, targetSheet.getMaxRows());
    const numRowsToProtect = lastRow - Config.DATA_START_ROW + 1;
    
    const p1 = targetSheet.getRange(Config.DATA_START_ROW, 1, numRowsToProtect, 2).protect().setDescription('システム専用(番号・氏名)');
    p1.removeEditors(p1.getEditors()); p1.addEditor(teacherEmail);
    
    const colsToProtectRight = Config.COL_MAIL_STATUS - Config.COL_FIX + 1;
    const p2 = targetSheet.getRange(Config.DATA_START_ROW, Config.COL_FIX, numRowsToProtect, colsToProtectRight).protect().setDescription('システム専用(添削結果)');
    p2.removeEditors(p2.getEditors()); p2.addEditor(teacherEmail);

    const configSheet = ss.getSheetByName(Config.API_KEY_SHEET);
    if (configSheet) configSheet.getRange(Config.LINK_SWITCH_CELL).clearDataValidations().setValue('🔒 ロック作動中（安全）'); 
    
    ss.toast('セットアップ完了', '✨ 完了', 5);
  },

  removeProtections: function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    const targetSheet = ss.getSheetByName(Config.MAIN_SHEET); 
    if (!targetSheet) return;
    if (targetSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).length === 0) return;

    const response = ui.alert('確認', '「添削」シートのすべての保護（ロック）を解除しますか？\n（生徒がどこでも編集できるようになります）', ui.ButtonSet.YES_NO);
    if (response !== ui.Button.YES) return;
    
    Security.clearProtections_(targetSheet); // ★分割に伴い Security に変更
    const configSheet = ss.getSheetByName(Config.API_KEY_SHEET);
    if (configSheet) configSheet.getRange(Config.LINK_SWITCH_CELL).clearDataValidations().setValue('🔓 解除中（生徒も編集可）'); 
    ss.toast('保護を解除しました。', '🔓 完了', 5);
  },

  clearProtections_: function(sheet) {
    const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    for (let i = 0; i < protections.length; i++) protections[i].remove();
  }
};