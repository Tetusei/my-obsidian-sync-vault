/**
 * 成績分析ツール (Grade Analysis Tool)
 * 
 * 目的: 学校のテスト結果（スプレッドシート）を自動で集計・分析し、クラスごとの推移を可視化する。
 * 使用環境: GAS (Google Apps Script), Gemini API
 */

function main() {
  console.log("成績分析ツールの処理を開始します。");
}

/**
 * スプレッドシートが開かれたときに自動で実行される関数（シンプルトリガー）
 * スプレッドシートのメニューバーにカスタムメニューを追加します。
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // 「📊 成績分析ツール」というカスタムメニューを作成します
  ui.createMenu("📊 成績分析ツール")
    .addItem("📈 クラス成績の分析を実行", "analyzeClassPerformance")
    .addSeparator()
    .addItem("🔧 APIキーの接続テストを実行", "testGeminiConnection")
    .addItem("🔍 使用可能なモデル一覧を取得", "listAvailableModels")
    .addToUi();
}

/**
 * スプレッドシートが編集されたときに自動で実行される関数（シンプルトリガー）
 * ユーザーがシート上のセルを編集した際に、この関数が自動的に呼び出されます。
 * 
 * @param {Object} e 編集イベントに関する情報を含むオブジェクト
 */
function onEdit(e) {
  // GASエディタから直接「実行」ボタンを押した場合など、引数 e が渡されない手動実行時のエラーを防止します
  if (!e || !e.range) {
    console.warn("イベントオブジェクト e が未定義のため、処理をスキップしました。この関数はスプレッドシートの編集時に自動で実行されます。");
    return;
  }

  // 1. 編集されたセル（範囲）とシートの情報を取得します
  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();
  
  // 2. 「設定」シートの「B2」セル（行:2, 列:2）が編集されたかを判定します
  // getRow() が 2 (2行目)、getColumn() が 2 (B列) の場合に対象とします
  if (sheetName === "設定" && range.getRow() === 2 && range.getColumn() === 2) {
    
    // 3. 入力された値（APIキー）を取得します
    const apiKey = range.getValue().toString().trim();
    
    // 値が空でない（有効な入力がある）場合のみ処理を行います
    if (apiKey !== "") {
      try {
        // 4. GASのスクリプトプロパティ（プロジェクト固有の隠し環境変数）を取得します
        const scriptProperties = PropertiesService.getScriptProperties();
        
        // 5. APIキーを「GEMINI_API_KEY」という名前で安全に保存します
        scriptProperties.setProperty("GEMINI_API_KEY", apiKey);
        
        // 6. セキュリティ対策のため、スプレッドシート上のB2セルに入力された文字列を即座に消去します
        // 値を空文字に設定し、さらにコンテンツをクリアします
        range.setValue("");
        range.clearContent();
        
        // 7. スプレッドシートの変更を強制的に即時適用（コミット）させます
        // これを呼ばないと、画面上の表示更新が遅れたり、反映されないことがあります
        SpreadsheetApp.flush();
        
        // 8. 保存が完了したことをユーザーに知らせるため、スプレッドシートの画面右下にトースト通知を表示します
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        // toast(表示メッセージ, タイトル, 表示時間（秒）)
        spreadsheet.toast("APIキーを安全に保存しました。", "保存完了", 5);
        
        console.log("APIキーが正常にスクリプトプロパティへ保存され、セルがクリアされました。");
        
      } catch (error) {
        // 例外が発生した場合はログに出力し、ユーザーにもエラーを通知します
        console.error("APIキーの保存中にエラーが発生しました: " + error.toString());
        SpreadsheetApp.getActiveSpreadsheet().toast(
          "APIキーの保存中にエラーが発生しました。詳細は実行ログを確認してください。", 
          "保存エラー", 
          10
        );
      }
    }
  }
}

/**
 * クラスごとのパフォーマンス（国語・算数・理科の各平均点および3教科合計の平均点）を分析し、
 * Gemini API からのAI指導アドバイス、および各科目の平均点比較グラフを含めて「分析結果」シートに書き出します。
 */
function analyzeClassPerformance() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. スクリプトプロパティから API キーを取得します
  const scriptProperties = PropertiesService.getScriptProperties();
  const apiKey = scriptProperties.getProperty("GEMINI_API_KEY");
  
  // APIキーが保存されていない場合は警告を出して処理を中断します
  if (!apiKey) {
    spreadsheet.toast(
      "GEMINI_API_KEY が設定されていません。「設定」シートのB2セルにAPIキーを入力してください。", 
      "設定エラー", 
      10
    );
    return;
  }
  
  // 2. 「データ」シートからデータを取得します
  const dataSheet = spreadsheet.getSheetByName("データ");
  if (!dataSheet) {
    spreadsheet.toast("「データ」シートが見つかりません。作成してください。", "エラー", 5);
    return;
  }
  
  // シートに書き込まれている全データを二次元配列として取得します
  const values = dataSheet.getDataRange().getValues();
  if (values.length <= 1) {
    spreadsheet.toast("データシートに入力されている成績データがありません（見出しのみ、または空です）。", "エラー", 5);
    return;
  }
  
  // 見出し行と実際のデータ行を分離します
  const headers = values[0]; // 1行目: 見出し（日付、クラス、氏名、国語、算数、理科、社会）
  const dataRows = values.slice(1); // 2行目以降の成績データ
  
  // 2. 見出しの並び順が変わっても正しく動作するよう、各項目の列番号（インデックス）を自動検出します
  const classIdx = headers.indexOf("クラス");
  const kokugoIdx = headers.indexOf("国語");
  const sansuIdx = headers.indexOf("算数");
  const rikaIdx = headers.indexOf("理科");
  
  // 必要な列が見つからない場合はエラーを表示して処理を終了します
  if (classIdx === -1 || kokugoIdx === -1 || sansuIdx === -1 || rikaIdx === -1) {
    spreadsheet.toast("必要な列（クラス、国語、算数、理科）が見つかりません。見出し名を確認してください。", "エラー", 7);
    return;
  }
  
  // 3. クラスごとに「国・算・理」の個別点、合計点、および生徒数を集計します
  // 集計用オブジェクトの構造: { "A組": { kokugoSum: 160, sansuSum: 120, rikaSum: 180, totalScore: 460, count: 2 }, ... }
  const classStats = {};
  
  dataRows.forEach(row => {
    const className = row[classIdx];
    // 各点数を数値に変換します（空欄などの場合は0点として扱います）
    const kokugo = Number(row[kokugoIdx]) || 0;
    const sansu = Number(row[sansuIdx]) || 0;
    const rika = Number(row[rikaIdx]) || 0;
    
    // 国語・算数・理科の3教科合計点
    const threeSubjectTotal = kokugo + sansu + rika;
    
    // クラス名が入力されている行のみ集計します
    if (className) {
      if (!classStats[className]) {
        classStats[className] = {
          kokugoSum: 0,
          sansuSum: 0,
          rikaSum: 0,
          totalScore: 0,
          count: 0
        };
      }
      classStats[className].kokugoSum += kokugo;
      classStats[className].sansuSum += sansu;
      classStats[className].rikaSum += rika;
      classStats[className].totalScore += threeSubjectTotal;
      classStats[className].count += 1;
    }
  });
  
  // 4. 分析の開始前に、Gemini APIのウォームアップ（コールドスタート対策）を実行します
  spreadsheet.toast("AIデータ分析エンジンを起動しています...", "準備中", 5);
  warmupGemini(apiKey);
  Utilities.sleep(1000); // 活性化後に1秒間待機します
  
  // 5. クラスごとの平均点を計算し、Gemini APIを呼び出してアドバイスを生成し、出力用の二次元配列を組み立てます
  // G列とH列は「グラフの描画スペース」として空けておき、I列にアドバイスを出力するよう構成します（デザインの最適化）
  const outputData = [["クラス", "国語平均", "算数平均", "理科平均", "3教科合計平均", "対象人数", "", "", "AI分析アドバイス"]]; 
  
  // クラス名をアルファベット・五十音順に並び替えて処理します
  const classNames = Object.keys(classStats).sort();
  classNames.forEach((className, idx) => {
    const stats = classStats[className];
    const count = stats.count;
    
    // 各科目の平均値の計算（生徒数が0名の場合は0点）
    const kokugoAvg = count > 0 ? Math.round((stats.kokugoSum / count) * 10) / 10 : 0;
    const sansuAvg = count > 0 ? Math.round((stats.sansuSum / count) * 10) / 10 : 0;
    const rikaAvg = count > 0 ? Math.round((stats.rikaSum / count) * 10) / 10 : 0;
    const totalAvg = count > 0 ? Math.round((stats.totalScore / count) * 10) / 10 : 0;
    
    // 2つ目以降のクラス処理の前に、APIの頻度制限（Rate Limit）を回避するため1.5秒の待機時間を設けます
    if (idx > 0) {
      Utilities.sleep(1500);
    }
    
    // 画面右下に現在のAI分析の進捗（クラス名など）をトースト表示します
    spreadsheet.toast(
      `クラス「${className}」のAI指導アドバイスを作成中（${idx + 1}/${classNames.length}クラス）...`, 
      "AI分析実行中", 
      15
    );
    
    // Gemini APIを呼び出して、各教科の平均データを踏まえたより具体的なアドバイスを生成します
    const advice = generateGeminiAdvice(apiKey, className, kokugoAvg, sansuAvg, rikaAvg, totalAvg, count);
    
    // 出力データに行を追加します（インデックス6, 7の空文字列はG列・H列のグラフ用スペースです）
    outputData.push([className, kokugoAvg, sansuAvg, rikaAvg, totalAvg, count, "", "", advice]);
  });
  
  // 6. 書き出し先となる「分析結果」シートを用意します
  let resultSheet = spreadsheet.getSheetByName("分析結果");
  if (resultSheet) {
    // 既にシートが存在する場合は、中身をクリアして初期化します
    resultSheet.clear();
  } else {
    // シートが存在しない場合は、新しく作成（挿入）します
    resultSheet = spreadsheet.insertSheet("分析結果");
  }
  
  // 7. 「分析結果」シートにデータを書き出します（9列分の領域）
  const outputRange = resultSheet.getRange(1, 1, outputData.length, 9);
  outputRange.setValues(outputData);
  
  // 8. 出力した表のデザイン（見た目）を綺麗に整えます
  // 全体の縦位置を中央揃えに設定
  outputRange.setVerticalAlignment("middle");
  
  // 数値データ部分（A〜F列: 1〜6列目）の罫線と配置の装飾
  const dataRange = resultSheet.getRange(1, 1, outputData.length, 6);
  dataRange
    .setBorder(true, true, true, true, true, true, "#e5e7eb", SpreadsheetApp.BorderStyle.SOLID)
    .setHorizontalAlignment("center");
  
  // AIアドバイス部分（I列: 9列目）の罫線、配置、およびテキスト折り返し装飾
  const adviceRange = resultSheet.getRange(1, 9, outputData.length, 1);
  adviceRange
    .setBorder(true, true, true, true, true, true, "#e5e7eb", SpreadsheetApp.BorderStyle.SOLID)
    .setHorizontalAlignment("left")
    .setWrap(true);
  
  // ヘッダー行（A〜F列、およびI列の1行目）のカラー装飾
  resultSheet.getRange(1, 1, 1, 6).setBackground("#3b82f6").setFontColor("#ffffff").setFontWeight("bold");
  resultSheet.getRange(1, 9, 1, 1).setBackground("#3b82f6").setFontColor("#ffffff").setFontWeight("bold");
  
  // 列幅を調整（データ部は自動、グラフスペースは固定幅、アドバイス列は長文用に広めの固定幅）
  resultSheet.autoResizeColumns(1, 6);
  resultSheet.setColumnWidth(7, 220); // G列（グラフスペース左）
  resultSheet.setColumnWidth(8, 220); // H列（グラフスペース右）
  resultSheet.setColumnWidth(9, 450); // I列（AIアドバイス）
  
  // 9. グラフのリフレッシュと新規作成処理
  // 既存のグラフがあれば二重作成を防ぐためにすべて削除します
  const charts = resultSheet.getCharts();
  charts.forEach(chart => {
    resultSheet.removeChart(chart);
  });
  
  // グラフの参照データ範囲（A1セル〜D列[最終行] = クラス名、国語、算数、理科）を指定します
  const chartDataRange = resultSheet.getRange(1, 1, outputData.length, 4);
  
  // 縦棒グラフを構築します
  const barChart = resultSheet.newChart()
    .setChartType(Charts.ChartType.COLUMN) // 縦棒グラフを指定
    .addRange(chartDataRange)
    .setPosition(2, 7, 0, 0) // G2セル（2行目、7列目）を基準に配置します
    .setOption("title", "クラス別 平均点比較")
    .setOption("hAxis", { title: "クラス" })
    .setOption("vAxis", { title: "平均点（100点満点）", minValue: 0, maxValue: 100 })
    .setOption("width", 430)  // G列・H列の幅（合計440px）に収まるサイズに調整
    .setOption("height", 280)
    .build();
  
  // 作成したグラフをシートに挿入します
  resultSheet.insertChart(barChart);
  
  // 10. 処理が完了したことをユーザーにトーストで知らせます
  spreadsheet.toast("AI分析および比較グラフの作成が完了しました！", "分析・描画完了", 5);
}

/**
 * Gemini API の接続を事前に温める（コールドスタート対策）ためのウォームアップ処理です。
 * 1回目のリクエストのタイムアウトや遅延による503エラーを軽減します。
 * 
 * @param {string} apiKey Gemini API キー
 */
function warmupGemini(apiKey) {
  const url = "https://generativelanguage.googleapis.com/v1/models/gemini-2.5-flash:generateContent?key=" + apiKey;
  const payload = {
    contents: [{
      parts: [{
        text: "ping"
      }]
    }]
  };
  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    UrlFetchApp.fetch(url, options);
    console.log("Gemini API ウォームアップ完了");
  } catch (e) {
    console.warn("ウォームアップ中に例外が発生しました（無視して進行します）: " + e.toString());
  }
}

/**
 * Gemini API を使用して、成績データに対するAI分析アドバイス（150文字程度）を生成します。
 * 各教科の個別平均点を考慮した、より高精度な指導アドバイスを生成します。
 * 
 * @param {string} apiKey Gemini API キー
 * @param {string} className 対象のクラス名
 * @param {number} kokugoAvg 国語の平均点
 * @param {number} sansuAvg 算数の平均点
 * @param {number} rikaAvg 理科の平均点
 * @param {number} totalAvg 3教科合計の平均点
 * @param {number} studentCount 対象の生徒数
 * @return {string} 生成されたアドバイス文章
 */
function generateGeminiAdvice(apiKey, className, kokugoAvg, sansuAvg, rikaAvg, totalAvg, studentCount) {
  const url = "https://generativelanguage.googleapis.com/v1/models/gemini-2.5-flash:generateContent?key=" + apiKey;

  // AIに渡す指示（プロンプト）を組み立てます
  const prompt = `あなたは小学校の優秀なデータアナリストです。
以下のクラスのテスト平均点データ（国語・算数・理科の各科目平均点および3教科合計の平均点）を分析し、このクラスの強みと、今後の学習指導における具体的なアドバイスを150文字程度で優しく日本語で作成してください。

【対象クラスのデータ】
クラス名: ${className}
対象人数: ${studentCount}名
国語の平均点: ${kokugoAvg}点 / 100点満点
算数の平均点: ${sansuAvg}点 / 100点満点
理科の平均点: ${rikaAvg}点 / 100点満点
3教科合計の平均点: ${totalAvg}点 / 300点満点`;

  // APIリクエスト用のボディを設定します
  const payload = {
    contents: [{
      parts: [{
        text: prompt
      }]
    }]
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const maxRetries = 5;  // 503エラーなどの一時的な不可に備え、最大5回リトライ
  let sleepTime = 2000;  // 最初のリトライ待機時間（2秒）

  for (let attempt = 0; attempt < maxRetries; attempt++) {
    try {
      // APIを呼び出します
      const response = UrlFetchApp.fetch(url, options);
      const responseCode = response.getResponseCode();
      const responseBody = response.getContentText();

      // 接続成功（HTTP 200）時
      if (responseCode === 200) {
        const json = JSON.parse(responseBody);
        
        // レスポンス構造からテキストコンテンツを抽出します
        if (json.candidates && json.candidates[0] && json.candidates[0].content && json.candidates[0].content.parts[0]) {
          return json.candidates[0].content.parts[0].text.trim();
        } else {
          return "（AIからの適切な回答が得られませんでした）";
        }
      } 
      
      // 接続失敗時
      console.warn(`Gemini API 接続失敗（試行 ${attempt + 1}/${maxRetries}回目、ステータス: ${responseCode}）: ${responseBody}`);
      
      // レート制限（429）やサーバー側一時的エラー（503などの500以上）の場合、少し時間を置いてリトライを試みます
      if (responseCode === 429 || responseCode >= 500) {
        if (attempt < maxRetries - 1) {
          Utilities.sleep(sleepTime);
          sleepTime *= 1.5; // 待機時間を徐々に延ばします（バックオフ）
          continue;
        }
      }
      
      return `（APIエラーのためアドバイスを生成できませんでした。コード: ${responseCode}）`;

    } catch (error) {
      console.error(`Gemini API 呼び出し中に例外が発生しました（試行 ${attempt + 1}/${maxRetries}回目）: ${error.toString()}`);
      
      if (attempt < maxRetries - 1) {
        Utilities.sleep(sleepTime);
        continue;
      }
      return "（通信エラーのためアドバイスを生成できませんでした。ネットワーク状況をご確認ください）";
    }
  }
  
  return "（アドバイスの生成に失敗しました）";
}

/**
 * Gemini APIとの通信テストを行うデバッグ用関数です。
 * エラーの詳細を画面（スプレッドシート上）にダイアログで表示します。
 */
function testGeminiConnection() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const scriptProperties = PropertiesService.getScriptProperties();
  const apiKey = scriptProperties.getProperty("GEMINI_API_KEY");
  
  if (!apiKey) {
    SpreadsheetApp.getUi().alert(
      "APIキー未設定", 
      "GEMINI_API_KEY がスクリプトプロパティに保存されていません。「設定」シートのB2セルにAPIキーを入力してください。", 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  
  // 安定版である gemini-2.5-flash を指定します
  const url = "https://generativelanguage.googleapis.com/v1/models/gemini-2.5-flash:generateContent?key=" + apiKey;
  const payload = {
    contents: [{
      parts: [{
        text: "Hello"
      }]
    }]
  };
  
  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();
    
    if (responseCode === 200) {
      SpreadsheetApp.getUi().alert(
        "接続成功！🎉", 
        "Gemini API と正常に通信できました。APIキーは有効です。", 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } else {
      // エラーの詳細を表示
      let errorMessage = `API接続に失敗しました。\n\n【HTTPステータスコード】\n${responseCode}\n\n【エラーの本体】\n${responseBody}`;
      
      // APIキーが無効である場合のヒント
      if (responseBody.includes("API_KEY_INVALID") || responseBody.includes("API key not valid")) {
        errorMessage += "\n\n💡 APIキーが間違っているか、正しくコピーできていない可能性があります。Google AI Studio で作成した正しいキーであるか再確認してください。";
      }
      // クォータ（上限）制限エラーの場合のヒント
      if (responseBody.includes("RESOURCE_EXHAUSTED")) {
        errorMessage += "\n\n💡 APIの利用制限（リクエスト上限）に達しています。しばらく時間をおいてから再度お試しください。";
      }
      
      SpreadsheetApp.getUi().alert("API接続エラー", errorMessage, SpreadsheetApp.getUi().ButtonSet.OK);
    }
  } catch (error) {
    SpreadsheetApp.getUi().alert(
      "例外エラー", 
      "通信自体に失敗しました。インターネット接続やGASの外部アクセス制限をご確認ください。\n\n詳細: " + error.toString(), 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * APIキーで使用可能なGeminiモデルの一覧を取得し、ダイアログで表示するデバッグ関数です。
 */
function listAvailableModels() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const scriptProperties = PropertiesService.getScriptProperties();
  const apiKey = scriptProperties.getProperty("GEMINI_API_KEY");
  
  if (!apiKey) {
    SpreadsheetApp.getUi().alert(
      "APIキー未設定", 
      "APIキーが設定されていません。「設定」シートで設定してください。", 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  
  const url = "https://generativelanguage.googleapis.com/v1/models?key=" + apiKey;
  
  try {
    const response = UrlFetchApp.fetch(url, { method: "get", muteHttpExceptions: true });
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();
    
    if (responseCode === 200) {
      const json = JSON.parse(responseBody);
      // モデル名を取得して整理します
      const modelNames = json.models ? json.models.map(m => m.name.replace("models/", "")) : [];
      let message = "このAPIキーで使用可能なモデル一覧:\n\n" + modelNames.join("\n");
      
      // アラートの表示制限に配慮して長さをトリミングします
      if (message.length > 1500) {
        message = message.substring(0, 1500) + "\n... (以下省略)";
      }
      
      SpreadsheetApp.getUi().alert("モデル一覧取得成功 🎉", message, SpreadsheetApp.getUi().ButtonSet.OK);
    } else {
      SpreadsheetApp.getUi().alert(
        "取得エラー", 
        `モデル一覧の取得に失敗しました。\n\n【ステータスコード】\n${responseCode}\n\n【エラー内容】\n${responseBody}`, 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  } catch (error) {
    SpreadsheetApp.getUi().alert(
      "例外エラー", 
      "通信時にエラーが発生しました。\n\n詳細: " + error.toString(), 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}
