/**
 * 【ファイル7】Handover.gs (v2.9.2)
 * 役割：AIへの引継ぎ用プロンプト（システム設計図）をドキュメントとして出力する
 * 構文エラーを修正した確定版
 */

function createHandoverDoc() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 保存先フォルダ設定
  let targetFolder;
  try {
    const ssFile = DriveApp.getFileById(ss.getId());
    const parentFolders = ssFile.getParents();
    targetFolder = parentFolders.hasNext() ? parentFolders.next() : DriveApp.getRootFolder();
  } catch(e) {
    targetFolder = DriveApp.getRootFolder(); 
  }
  
  // ファイル名設定
  const dateStr = Utilities.formatDate(new Date(), "JST", "yyyyMMdd_HHmm");
  const version = (typeof CONFIG !== 'undefined' && CONFIG.VERSION) ? CONFIG.VERSION : 'v2.9.2';
  const docName = "【AI引継用】システム構成書_" + version + "_" + dateStr;
  
  // ドキュメント作成
  const doc = DocumentApp.create(docName);
  const docFile = DriveApp.getFileById(doc.getId());
  docFile.moveTo(targetFolder); 

  // ■ AIに読ませるシステム設計図（プロンプト）
  const promptText = `
あなたはGoogle Apps Script (GAS) と Gemini API のエキスパートです。
現在運用中の「学校給食アレルギー管理システム (${version})」の保守・開発を引き継ぎます。

【★開発・修正時の絶対ルール】

1. 日付の「OCR優先」原則 (The Year-2027 Trap回避)
   - 「PDF紙面の年号」を絶対正義とする。令和8年は現在に関わらず2026年として処理。

2. 安全第一のデータ処理
   - 処理済みPDFは即座に移動。カレンダー同期・通知は B8(Botスイッチ) ONの時のみ実行。

3. 設定参照の徹底
   - シート名やセル番地はハードコードせず、必ず Config.gs の定義を経由すること。

4. 判定ロジックの「合体・部分一致」原則 (松尾・中尾対策)
   - 生徒側：(アレルゲン+料理名+対応内容) を結合し、記号除去後のテキストで部分一致判定を行う。表記揺れを許容し、判定漏れをゼロにすること。

5. B10セルの「毎日」限定フィルター仕様
   - B10の除外ワードは、日付が「毎日」の生徒にのみ適用。
   - 特定日付（個別対応）の生徒は、安全のためB10に関わらず【必ず表示】させること。

6. 辞書の自動学習・バックアップ機能 (onEdit)
   - VERIFYシートH列への「【分類：単語】」入力検知時、自動で辞書末尾に追加。
   - 登録と同時に「Backup_Dict_yyyyMMdd」形式でシートをコピー保存し、トースト通知を行う。

7. エラーハンドリングの徹底
   - setRichTextValue 等の実行時は、空文字やnullによる「引数が無効です」エラーを回避するため、必ず長さチェックと型変換を行うこと。

8. 対話および実装フェーズの厳守
   - ユーザーからの明確な「コード」指示があるまで、GASを出力しない。
   - モデル名は 'gemini-flash-latest' 固定。省略のない完全なコードを提供すること。
   - ユーザーの「正本」にあるコメントや命名規則をAIの判断で書き換えることを厳禁とする。
`;

  doc.getBody().setText(promptText);

  // 完了メッセージの表示（構文エラー回避のため文字列結合で記述）
  const htmlText = 
    '<div style="font-family:sans-serif; text-align:center;">' +
      '<p>✅ 完璧な引継ぎ書（v2.9.2）を作成しました。</p>' +
      '<p style="margin:10px 0;">ファイル名: <strong>' + docName + '</strong></p>' +
      '<a href="' + doc.getUrl() + '" target="_blank" style="' +
        'display:inline-block;' +
        'background-color:#0b8043;' +
        'color:#fff;' +
        'padding:10px 20px;' +
        'text-decoration:none;' +
        'border-radius:4px;' +
        'font-weight:bold;">📄 ドキュメントを開く</a>' +
    '</div>';
  
  const html = HtmlService.createHtmlOutput(htmlText).setWidth(400).setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(html, '作成完了');
}