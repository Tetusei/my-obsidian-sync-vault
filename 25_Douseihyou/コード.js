// スプレッドシートを開いた時に自動でメニューを追加する関数
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('★校務便利機能')
    .addItem('シートURL目次とシート名の変更', 'renameAndGenerateLinks')
    .addItem('各シート氏名欄の一括変更', 'distributeNamesAndJobs_v3')
    .addItem('個人別シートの一括コピー', 'copyBodyAndRowHeights')
    .addItem('各シートのデータクリア', 'clearRangeAndX5')
    .addToUi();
}

// 職員動静一覧のF3セルから右への名前書き込み、個人別シート名変更、基礎データの目次リンク生成を行う関数
function renameAndGenerateLinks() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. 「基礎データ」シートを取得
  var targetSheetName = "基礎データ";
  var targetSheet = ss.getSheetByName(targetSheetName);
  if (!targetSheet) {
    Browser.msgBox("エラー: '" + targetSheetName + "' シートが見つかりません。");
    return;
  }
  
  // 2. 「職員動静一覧」シートを取得
  var listSheetName = "職員動静一覧";
  var listSheet = ss.getSheetByName(listSheetName);
  if (!listSheet) {
    Browser.msgBox("エラー: '" + listSheetName + "' シートが見つかりません。");
    return;
  }

  // 3. B列（氏名）から名前リストを取得（2行目以降）
  var lastRow = targetSheet.getLastRow();
  if (lastRow < 2) {
    Browser.msgBox("対象データがありません。「基礎データ」シートのB列を確認してください。");
    return;
  }
  
  var namesData = targetSheet.getRange(2, 2, lastRow - 1, 1).getValues();
  var names = [];
  for (var i = 0; i < namesData.length; i++) {
    var val = namesData[i][0].toString().trim();
    if (val !== "") {
      names.push(val);
    }
  }

  if (names.length === 0) {
    Browser.msgBox("「基礎データ」シートのB列に名前が入力されていません。");
    return;
  }

  // 4. 「職員動静一覧」の3行目・F列（Column 6）から右方向に基礎データのセル参照（数式）を入力
  // ※ F3G3, H3I3...のように2列ずつ結合（マージ）されているため、1つおきに空文字を挿入します。
  var lastCol = listSheet.getLastColumn();
  if (lastCol >= 6) {
    listSheet.getRange(3, 6, 1, lastCol - 5).clearContent();
  }
  var formulaRow = [];
  for (var k = 0; k < names.length; k++) {
    formulaRow.push("=基礎データ!B" + (k + 2));
    formulaRow.push(""); // 結合セルの右側分
  }
  if (formulaRow.length > 0) {
    listSheet.getRange(3, 6, 1, formulaRow.length).setFormulas([formulaRow]);
  }

  // --- 追加機能：6行目から最後の行までの範囲に VLOOKUP 数式を自動書き込み ---
  var listLastRow = listSheet.getLastRow();
  var numRows = listLastRow - 6 + 1;
  if (numRows > 0) {
    // 既存の6行目以降のF列から右端までの数式をクリア
    var clearWidth = Math.max(lastCol - 5, formulaRow.length);
    listSheet.getRange(6, 6, numRows, clearWidth).clearContent();
    
    var formulas2D = [];
    for (var r = 6; r <= listLastRow; r++) {
      var rowFormulas = [];
      for (var k = 0; k < names.length; k++) {
        var col = 6 + 2 * k;
        var colLetter = getColumnLetter(col);
        // =iferror(VLOOKUP(TRIM(INDIRECT("'" & $ColLetter$3 & "'!E" + (r + 9))), '基礎データ'!$F$3:$G$18, 2, FALSE),"")
        var formula = '=iferror(VLOOKUP(TRIM(INDIRECT("\'" & $' + colLetter + '$3 & "\'!E' + (r + 9) + '")), \'基礎データ\'!$F$3:$G$18, 2, FALSE),"")';
        rowFormulas.push(formula);
        rowFormulas.push(""); // 結合セルの右側分
      }
      formulas2D.push(rowFormulas);
    }
    
    if (formulas2D.length > 0 && formulas2D[0].length > 0) {
      listSheet.getRange(6, 6, numRows, formulas2D[0].length).setFormulas(formulas2D);
    }
  }

  // 5. 個人別シートの確認と自動作成（「基礎データ」「職員動静一覧」以外のシート）
  var allSheets = ss.getSheets();
  var sheetMap = {};
  var individualSheets = [];
  for (var i = 0; i < allSheets.length; i++) {
    var s = allSheets[i];
    var sName = s.getName();
    sheetMap[sName] = s;
    if (sName !== "基礎データ" && sName !== "職員動静一覧") {
      individualSheets.push(s);
    }
  }

  // 最初の個人別シートをテンプレートとして使用する
  var templateSheet = individualSheets.length > 0 ? individualSheets[0] : null;

  // 不足しているシートがあれば自動でテンプレートから複製する
  var createdCount = 0;
  for (var k = 0; k < names.length; k++) {
    var name = names[k];
    var sanitizedName = name.replace(/\s+/g, '');
    
    if (!sheetMap[sanitizedName]) {
      if (templateSheet) {
        try {
          var newSheet = templateSheet.copyTo(ss);
          newSheet.setName(sanitizedName);
          sheetMap[sanitizedName] = newSheet;
          individualSheets.push(newSheet);
          createdCount++;
        } catch (e) {
          console.log("シート複製エラー (" + sanitizedName + "): " + e.message);
        }
      }
    }
  }

  // 6. 各名前について、最新のシートURLを取得しハイパーリンクを作成
  var basicUrl = ss.getUrl();
  var listForB = [];
  var listForD = [];

  for (var k = 0; k < names.length; k++) {
    var name = names[k];
    var sanitizedName = name.replace(/\s+/g, '');
    var sheet = sheetMap[sanitizedName];
    
    if (sheet) {
      var sheetId = sheet.getSheetId();
      var sheetUrl = basicUrl + "#gid=" + sheetId;
      var safeName = name.replace(/"/g, '""');
      var formula = '=HYPERLINK("' + sheetUrl + '", "' + safeName + '")';
      listForB.push([formula]);
      listForD.push([sheetUrl]);
    } else {
      listForB.push([name]);
      listForD.push([""]);
    }
  }

  // 7. 「基礎データ」シートへの書き込み
  // names.length より後ろにある古いデータをクリア
  var maxRows = targetSheet.getMaxRows();
  if (maxRows > names.length + 1) {
    var extraRows = maxRows - (names.length + 1);
    targetSheet.getRange(names.length + 2, 2, extraRows, 1).clearContent();
    targetSheet.getRange(names.length + 2, 4, extraRows, 1).clearContent();
  }

  if (names.length > 0) {
    targetSheet.getRange(2, 2, names.length, 1).setFormulas(listForB);
    targetSheet.getRange(2, 4, names.length, 1).setValues(listForD);
  }

  // 8. シートの並び順（タブ順）の再配置
  var originalActiveSheet = ss.getActiveSheet();
  
  // 1枚目「基礎データ」、2枚目「職員動静一覧」を先頭に固定
  var kisoSheet = ss.getSheetByName("基礎データ");
  if (kisoSheet) {
    ss.setActiveSheet(kisoSheet);
    ss.moveActiveSheet(1);
  }
  var shokuinDouseiSheet = ss.getSheetByName("職員動静一覧");
  if (shokuinDouseiSheet) {
    ss.setActiveSheet(shokuinDouseiSheet);
    ss.moveActiveSheet(2);
  }
  
  // 3枚目以降に「基礎データ」のB列の並び順通りに個人別シートを並べる
  for (var k = 0; k < names.length; k++) {
    var name = names[k];
    var sanitizedName = name.replace(/\s+/g, '');
    var sheet = sheetMap[sanitizedName];
    if (sheet) {
      ss.setActiveSheet(sheet);
      ss.moveActiveSheet(3 + k);
    }
  }

  // 元のアクティブシートにフォーカスを戻す（無効な場合は「基礎データ」）
  try {
    ss.setActiveSheet(originalActiveSheet);
  } catch (e) {
    if (kisoSheet) {
      ss.setActiveSheet(kisoSheet);
    }
  }

  var msg = "完了しました！\n" +
            "・読み込んだ氏名の数: " + names.length + "人\n" +
            "・個人別シートの総数: " + individualSheets.length + "枚 (新規作成: " + createdCount + "枚)\n\n" +
            "※「職員動静一覧」の結合セル構造に合わせた数式展開、シートの自動複製、シートの並び順（タブ順）の並び替え、および基礎データへのリンク生成が完了しました。";
  Browser.msgBox(msg);
}

// 各個人別シートの氏名や職種を割り振る関数
function distributeNamesAndJobs_v3() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  // 1. 「基礎データ」シートを取得
  const sourceSheet = ss.getSheetByName("基礎データ");
  if (!sourceSheet) {
    Browser.msgBox("エラー: 「基礎データ」というシートが見つかりません。");
    return;
  }

  // 2. データの準備
  // 固定用：B2セルの値（役職者名）を取得
  const fixedName = sourceSheet.getRange("B2").getValue();

  const lastRow = sourceSheet.getLastRow();
  if (lastRow < 2) {
    Browser.msgBox("「基礎データ」シートにデータがありません。");
    return;
  }

  // B列(氏名)とC列(職種)を取得
  const dataList = sourceSheet.getRange(2, 2, lastRow - 1, 2).getValues();

  // 個人別シートを特定（「基礎データ」「職員動静一覧」以外）
  const targetSheets = [];
  for (let i = 0; i < sheets.length; i++) {
    let sName = sheets[i].getName();
    if (sName !== "基礎データ" && sName !== "職員動静一覧") {
      targetSheets.push(sheets[i]);
    }
  }

  // 3. 転記処理
  let listIndex = 0; 

  for (let i = 0; i < targetSheets.length; i++) {
    let targetSheet = targetSheets[i];

    // 名簿の人数分が終わったら終了
    if (listIndex >= dataList.length) {
      break;
    }

    // A) D7セル: 常に「基礎データ!B2」の値（役職者）を入れる
    targetSheet.getRange("D7").setValue(fixedName);

    // B) U9セル: 名簿リストの氏名（B列の値）
    targetSheet.getRange("U9").setValue(dataList[listIndex][0]);

    // C) P9セル: 名簿リストの職種（C列の値）
    targetSheet.getRange("P9").setValue(dataList[listIndex][1]);

    listIndex++;
  }

  Browser.msgBox("完了しました！\n・D7に役職者名(固定)\n・U9に氏名(個別)\n・P9に職種(個別)\nを転記しました。");
}

// 最初の個人別シートの本文データを、その他のすべての個人別シートに複製する関数（行の高さも同期）
function copyBodyAndRowHeights() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  // 個人別シートを特定
  const individualSheets = [];
  for (let i = 0; i < sheets.length; i++) {
    let sName = sheets[i].getName();
    if (sName !== "基礎データ" && sName !== "職員動静一覧") {
      individualSheets.push(sheets[i]);
    }
  }

  if (individualSheets.length === 0) {
    Browser.msgBox("コピー対象となる個人別シートが見つかりません。");
    return;
  }

  // 「職員動静一覧」シートのURLを生成
  const listSheet = ss.getSheetByName("職員動静一覧");
  if (!listSheet) {
    Browser.msgBox("エラー: '職員動静一覧' シートが見つかりません。");
    return;
  }
  const listSheetUrl = ss.getUrl() + "#gid=" + listSheet.getSheetId();

  // 1. コピー元のシート（最初の個人別シート）
  const sourceSheet = individualSheets[0];
  const sourceLastRow = sourceSheet.getLastRow();

  // データチェック
  if (sourceLastRow < 12) {
    Browser.msgBox("コピー元のデータが足りません（11行目から最後の1つ前までが必要です）。");
    return;
  }

  // コピーする行数などを計算
  const startRow = 11;
  const numRowsToCopy = (sourceLastRow - 1) - startRow + 1;
  const sourceRange = sourceSheet.getRange(startRow, 1, numRowsToCopy, sourceSheet.getLastColumn());

  // 元の行の高さを1行ずつ記憶しておく
  let rowHeights = [];
  for (let r = 0; r < numRowsToCopy; r++) {
    rowHeights.push(sourceSheet.getRowHeight(startRow + r));
  }

  // すべての個人別シートのZ1セルにURLを書き込み、文字色を白にする
  for (let i = 0; i < individualSheets.length; i++) {
    let targetSheet = individualSheets[i];
    let z1Range = targetSheet.getRange("Z1");
    z1Range.setValue(listSheetUrl);
    z1Range.setFontColor("#ffffff");
  }

  // 2. 2番目の個人別シート以降すべてに対して処理
  for (let i = 1; i < individualSheets.length; i++) {
    let targetSheet = individualSheets[i];
    let targetLastRow = targetSheet.getLastRow();

    // A) 古い本文データを削除
    if (targetLastRow > startRow) {
      let rowsToDelete = (targetLastRow - 1) - startRow + 1;
      if (rowsToDelete > 0) {
        targetSheet.deleteRows(startRow, rowsToDelete);
      }
    }
    
    // B) 新しい行を挿入
    targetSheet.insertRows(startRow, numRowsToCopy);

    // C) データをコピー
    sourceRange.copyTo(targetSheet.getRange(startRow, 1));

    // D) 記憶しておいた「行の高さ」を1行ずつ適用する
    for (let r = 0; r < numRowsToCopy; r++) {
      targetSheet.setRowHeight(startRow + r, rowHeights[r]);
    }
  }

  Browser.msgBox("完了しました！\n本文のコピー、行の高さ調整、およびZ1セルへの職員動静一覧URL書き込み（白色）が完了しました。");
}

// 各個人別シートの入力枠をクリアする関数
function clearRangeAndX5() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  // 各個人別シート（基礎データ、職員動静一覧以外）に対して処理
  for (let i = 0; i < sheets.length; i++) {
    let sheet = sheets[i];
    let sName = sheet.getName();
    if (sName === "基礎データ" || sName === "職員動静一覧") {
      continue;
    }

    // A) X5セルをクリア
    sheet.getRange("X5").clearContent();

    // B) E14:K列の範囲クリア
    let lastRow = sheet.getLastRow();
    
    // データが15行以上ある場合のみ範囲クリアを実行（行14以降）
    if (lastRow >= 15) {
      let rowsToClear = (lastRow - 1) - 14 + 1;
      // E列(5)からK列までの7列分
      sheet.getRange(14, 5, rowsToClear, 7).clearContent();
    }

    // C) 連絡先セルのクリア
    sheet.getRange("D52:M52").clearContent();
    sheet.getRange("P52:Y52").clearContent();
  }

  Browser.msgBox("完了しました！\n全シートの X5セル、E14:K の範囲、および連絡先セル（D52:M52, P52:Y52）をクリアしました。");
}

// 列インデックス（1始まり）をアルファベット表現（A, B, C... AA, AB...）に変換するヘルパー関数
function getColumnLetter(col) {
  var letter = "";
  while (col > 0) {
    var temp = (col - 1) % 26;
    letter = String.fromCharCode(65 + temp) + letter;
    col = (col - temp - 1) / 26;
  }
  return letter;
}