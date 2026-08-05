/**
 * @file Admin.gs
 * @description カスタムメニューの作成および管理者向け・メンテナンス用機能を管理します。
 * @version v1.1.0
 */

/**
 * スプレッドシートを開いたときに実行される関数。
 * 上部のカスタムメニューに「📄PDF連携機能」を追加します。
 * （※Main.gsからこちらのAdmin.gsへ構造を分離・移行しました）
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📄PDF連携機能')
    .addItem('🚀 すべて一括実行', 'executeAllProcesses')
    .addSeparator()
    .addItem('① ファイル名生成 (K列)', 'generateFileNamesFromEmailAndName')
    .addItem('② 拡張子を大文字(PDF)に変換', 'renamePdfExtensions')
    .addItem('③ フォルダ一覧更新', 'getFileNamesAndUrlsFromFolder')
    .addItem('④ URLを一括書き込み(一覧シート)', 'setPdfUrls')
    .addItem('⑤ ファイル照合チェック', 'checkFileUrlMatches')
    .addSeparator()
    .addSubMenu(
      ui.createMenu('🛠️ メンテナンス')
        .addItem('📦 手動バックアップを作成', 'adminCreateManualBackup')
        .addItem('ℹ️ システムバージョン確認', 'adminShowVersionInfo')
    )
    .addToUi();
}

/**
 * 【管理者機能】メニューから手動でいつでもバックアップを作成できる関数
 */
function adminCreateManualBackup() {
  SpreadsheetApp.getActiveSpreadsheet().toast('手動バックアップを作成しています...', '実行中');
  try {
    createBackupOfCurrentFile();
    Browser.msgBox('✅ バックアップ完了\n「' + Config.BACKUP_FOLDER_NAME + '」フォルダに現在の状態を安全に保存しました。');
  } catch (e) {
    Browser.msgBox('エラー: バックアップの作成に失敗しました。\n' + e.message);
  }
}

/**
 * 【管理者機能】現在の稼働バージョンを画面上にポップアップ表示する関数
 */
function adminShowVersionInfo() {
  Browser.msgBox(
    '📄 システム情報',
    '【ICT支援員補佐 GASシステム】\n\n' +
    '現在稼働中のバージョン: ' + Config.VERSION + '\n' +
    '対象設定年度: 2026年度（令和8年）\n\n' +
    '学校現場の業務効率化を応援しています！',
    Browser.Buttons.OK
  );
}