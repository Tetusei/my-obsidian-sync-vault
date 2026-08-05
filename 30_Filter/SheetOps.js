/**
 * @file SheetOps.gs
 * @description 「R8年度」シートへのデータ自動追記とレイアウト制御を専門に行います。
 */

/**
 * 分析結果と生徒情報をスプレッドシート（R8年度シート）の最終行に追記する
 * @param {Date} messageDate メールの受信日時
 * @param {Object} studentInfo 名簿から一致した生徒情報オブジェクト
 * @param {string} email 生徒のメールアドレス
 * @param {string} reason 検知理由
 * @param {string} accessCount アクセス回数
 * @param {Object} aiResult Geminiによる分析結果
 */
function recordToLogSheet(messageDate, studentInfo, email, reason, accessCount, aiResult) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = Config.SHEET_NAME_RECORD_PREFIX; // 令和8年度シート名
  let sheet = ss.getSheetByName(sheetName);
  
  // もしシートがなければ自動で新規作成し、きれいな見出しを配置する
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    const headers = ["日時", "(日時)", "時間", "学年", "組", "氏 名", "アドレス", "内 容", "対 応", "備 考"];
    sheet.appendRow(headers);
    // 見出し行を少し綺麗に装飾（太字＋やさしいブルー）
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#e6f2ff");
  }

  // 2026年（令和8年）基準に準拠した日付と時間の切り分け処理
  const dateStr = Utilities.formatDate(messageDate, "JST", "yyyy/MM/dd");
  
  // 曜日を自動計算して「(水)」のような形式にする
  const weekdays = ["日", "月", "火", "水", "木", "金", "土"];
  const dayStr = `(${weekdays[messageDate.getDay()]})`;
  
  const timeStr = Utilities.formatDate(messageDate, "JST", "HH:mm");

  // 生徒情報の紐解き
  const grade = studentInfo ? studentInfo.grade : "";
  const className = studentInfo ? studentInfo.class : "";
  const name = studentInfo ? studentInfo.name : "⚠️ 名簿に該当なし";

  // 既存データレイアウトを尊重し、内容列に理由と回数を格納
  const contentStr = `理由: ${reason} (アクセス回数: ${accessCount}回)`;
  
  // 対応状況の初期設定
  const statusStr = "未対応";

  // 備考欄にGemini AIの高度な判定結果をすっきり集約
  const remarkStr = `【AI緊急度: ${aiResult.urgency}】 意図: ${aiResult.intention} ／ 要約: ${aiResult.summary}`;

  // レイアウト順に行データを構築
  const rowData = [
    dateStr,      // 日時 (A列)
    dayStr,       // (日時) 曜日 (B列)
    timeStr,      // 時間 (C列)
    grade,        // 学年 (D列)
    className,    // 組 (E列)
    name,         // 氏 名 (F列)
    email,        // アドレス (G列)
    contentStr,   // 内 容 (H列)
    statusStr,    // 対 応 (I列)
    remarkStr     // 備 考 (J列)
  ];

  // シートの一番下の行にサッと追記
  sheet.appendRow(rowData);
}