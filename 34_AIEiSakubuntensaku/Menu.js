/**
 * ファイル名: Menu.gs
 * 役割: カスタムメニューの作成
 * バージョン: v5.3.0 (セパレータ追加・UI改善版)
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const prepMenu = ui.createMenu('🏫 授業の準備・分析')
    .addItem('📝 生徒用提出フォームを自動作成', 'FormManager.createStudentForm')
    .addItem('⏳ フォームの自動締め切りを設定', 'FormManager.setFormDeadline') 
    .addSeparator() 
    .addItem('📊 成績を名簿シートに集計・同期', 'DataSync.syncScoresToRoster')
    .addItem('🔄 履歴の回数をリセットする', 'FormManager.resetHistoryCount') 
    .addSeparator()
    .addItem('📊 クラス全体の傾向を分析', 'Analysis.analyzeClass'); 
    
  const sysMenu = ui.createMenu('⚙️ システム管理・設定')
    .addItem('🔍 現在のロック（保護）状態を確認', 'Security.checkProtectionStatus')
    .addSeparator()
    .addItem('🔒 名簿連動セットアップ（保護を開始）', 'Security.setupRosterLink')
    .addItem('🔓 保護（ロック）の一括解除', 'Security.removeProtections')
    .addSeparator()
    .addItem('📦 スクリプトとデータのバックアップ作成', 'DataSync.createBackup');
    
  ui.createMenu('🤖 AI添削システム')
    .addItem('▶️ 1. 未処理の生徒を一括添削', 'Main.processUnprocessedRows')
    .addSeparator() // ★ ここに区切り線（セパレータ）を追加しました！
    .addItem('📩 2. 添削結果をメールで一斉送信', 'Mail.sendEmails') 
    .addItem('🔔 3. 未提出の生徒へリマインド送信', 'Mail.sendReminders') 
    .addSeparator() 
    .addSubMenu(prepMenu) 
    .addSubMenu(sysMenu)  
    .addToUi();
}