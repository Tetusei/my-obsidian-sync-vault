/**
 * ファイル名: Analysis.gs
 * 役割: クラス全体の傾向分析・レポート作成
 * バージョン: v5.3.0 (トースト表示自動消去対応)
 */

const Analysis = {
  analyzeClass: function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(Config.MAIN_SHEET);
    const ui = SpreadsheetApp.getUi();
    
    if (!sheet) { ui.alert(`エラー: 「${Config.MAIN_SHEET}」シートが見つかりません。`); return; }
    
    const lastRow = sheet.getLastRow();
    if (lastRow < Config.DATA_START_ROW) { ui.alert('生徒の英文が入力されていません。'); return; }
    
    const texts = sheet.getRange(Config.DATA_START_ROW, Config.COL_STUDENT, lastRow - Config.DATA_START_ROW + 1, 1).getValues().flat().filter(t => t);
    if (texts.length === 0) { ui.alert('生徒の英文が入力されていません。'); return; }

    const response = ui.alert('確認', `クラス全員（${texts.length}件）の英作文をAIに分析させますか？\n（※「${Config.MAIN_SHEET}」シートのデータが対象です）`, ui.ButtonSet.YES_NO);
    if (response !== ui.Button.YES) return;

    ss.toast('AIがクラス全体の傾向を分析しています...', '📊 分析中', -1);
    let apiKeys = []; try { apiKeys = getApiKeys(); } catch(e) { ui.alert(e.message); return; }
    
    const prompt = `あなたは優秀な中学校の英語教師です。以下の生徒の英作文を分析し、レポートを作成してください。\n\n# 📊 クラス全体の傾向と分析レポート\n## 1. よくある間違いと傾向\n## 2. 次回の授業で取り上げるべき文法ポイント\n## 3. 先生への指導アドバイス\n\n【リスト】\n${texts.join('\n')}`;

    const payload = { "contents": [{ "parts": [{ "text": prompt }] }] };
    const options = { "method": "post", "contentType": "application/json", "payload": JSON.stringify(payload), "muteHttpExceptions": true };
    
    let reportText = null; let errorMessage = "エラーが発生しました。";
    const maxRetries = Math.max(3, apiKeys.length * 2);

    for (let i = 0; i < maxRetries; i++) {
      const currentKey = apiKeys[i % apiKeys.length];
      const url = `https://generativelanguage.googleapis.com/v1beta/models/${Config.MODEL_NAME}:generateContent?key=${currentKey}`;
      try {
        const res = UrlFetchApp.fetch(url, options);
        if (res.getResponseCode() === 429) { Utilities.sleep(2000); continue; }
        const json = JSON.parse(res.getContentText());
        if (json.candidates) { reportText = json.candidates[0].content.parts[0].text; break; }
        else if (json.error) {
          if (json.error.message.includes("quota") || json.error.message.includes("429")) { Utilities.sleep(2000); continue; }
          errorMessage = "AIエラー: " + json.error.message; break;
        }
      } catch (e) {
        if (e.toString().includes("429") || e.toString().includes("Quota")) { Utilities.sleep(2000); continue; }
        errorMessage = "通信エラー: " + e.toString(); break;
      }
    }
    
    if (reportText) {
      let reportSheet = ss.getSheetByName(Config.REPORT_SHEET);
      if (!reportSheet) reportSheet = ss.insertSheet(Config.REPORT_SHEET);
      reportSheet.clear();
      reportSheet.getRange('A1').setValue(reportText).setWrap(true);
      reportSheet.setColumnWidth(1, 800);
      ss.setActiveSheet(reportSheet);
      // ★ 成功時はこれで上書きされて消えます
      ss.toast('分析レポートが完成しました！', '✨ 完了', 5);
    } else { 
      // ★ 追加: 失敗時も出っぱなしのトーストを上書きして消す
      ss.toast('分析を終了しました。', 'ℹ️ お知らせ', 3);
      ui.alert('分析エラー\n\n詳細: ' + errorMessage); 
    }
  }
};