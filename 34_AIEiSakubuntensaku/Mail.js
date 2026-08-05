/**
 * ファイル名: Mail.gs
 * 役割: 添削結果の一斉送信・未提出者へのリマインド送信
 * バージョン: v5.3.0 (トースト表示自動消去対応)
 */

const Mail = {
  sendEmails: function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(Config.MAIN_SHEET);
    const ui = SpreadsheetApp.getUi();
    const rosterSheet = ss.getSheetByName(Config.ROSTER_SHEET);

    if (!sheet) { ui.alert(`エラー: 「${Config.MAIN_SHEET}」シートが見つかりません。`); return; }
    if (!rosterSheet) { ui.alert('エラー: 名簿シートが見つかりません。'); return; }

    const response = ui.alert('確認', '添削結果を生徒に一斉送信します。\n（※すでに「送信済」の生徒には重複して送信されません）', ui.ButtonSet.YES_NO);
    if (response !== ui.Button.YES) return;

    ss.toast('メール送信を準備しています...', '📩 送信中', -1);

    const rosterData = rosterSheet.getDataRange().getValues();
    const emailMap = {};
    for (let i = 1; i < rosterData.length; i++) {
      if (rosterData[i][0] && rosterData[i][1]) emailMap[rosterData[i][0]] = rosterData[i][1];
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < Config.DATA_START_ROW) return;
    const data = sheet.getRange(Config.DATA_START_ROW, 1, lastRow - Config.DATA_START_ROW + 1, Config.COL_MAIL_STATUS).getValues();
    let sendCount = 0;

    for (let i = 0; i < data.length; i++) {
      const num = data[i][0];
      const studentText = data[i][Config.COL_STUDENT - 1];
      const fixText = data[i][Config.COL_FIX - 1];
      const explainText = data[i][Config.COL_EXPLAIN - 1];
      const scoreText = data[i][Config.COL_SCORE - 1];
      const mailStatus = data[i][Config.COL_MAIL_STATUS - 1]; 

      if (mailStatus && mailStatus.toString().includes('送信済')) continue;

      if (num && studentText && fixText && !fixText.includes('生成中') && !fixText.includes('エラー')) {
        const emailAddress = emailMap[num];
        if (emailAddress) {
          const subject = "【英語科より】AI英作文 添削結果のお知らせ";
          const plainBody = `お疲れ様です！あなたの英作文のフィードバックをお届けします。\n\n【あなたの書いた英文】\n${studentText}\n\n【添削結果・ヒント】\n${fixText}\n\n【解説】\n${explainText}\n\n【スコア】\n${scoreText ? scoreText : '-'}`;
          const htmlBody = `
            <div style="font-family: 'Helvetica Neue', Arial, sans-serif; color: #333; max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">
              <div style="background-color: #4285F4; color: white; padding: 20px; text-align: center;">
                <h2 style="margin: 0; font-size: 22px;">✨ AI英作文 フィードバック ✨</h2>
              </div>
              <div style="padding: 25px; background-color: #ffffff;">
                <p style="font-size: 15px; margin-bottom: 25px;">お疲れ様です！あなたの英作文の添削結果が届きました。</p>
                <div style="margin-bottom: 25px;">
                  <h3 style="color: #4285F4; font-size: 16px; margin-bottom: 8px;">📝 あなたの書いた英文</h3>
                  <div style="background-color: #f8f9fa; padding: 15px; border-left: 4px solid #4285F4; border-radius: 4px; font-size: 16px;">${studentText}</div>
                </div>
                <div style="margin-bottom: 25px;">
                  <h3 style="color: #EA4335; font-size: 16px; margin-bottom: 8px;">💡 添削結果・ヒント</h3>
                  <div style="background-color: #fce8e6; padding: 15px; border-left: 4px solid #EA4335; border-radius: 4px; font-size: 16px; font-weight: bold;">${fixText}</div>
                </div>
                <div style="margin-bottom: 25px;">
                  <h3 style="color: #34A853; font-size: 16px; margin-bottom: 8px;">📖 先生からの解説</h3>
                  <div style="background-color: #e6f4ea; padding: 15px; border-left: 4px solid #34A853; border-radius: 4px; font-size: 15px; line-height: 1.5;">${explainText}</div>
                </div>
                <div style="text-align: center; margin-top: 35px; margin-bottom: 10px;">
                  <span style="display: inline-block; background-color: #FBBC04; color: #ffffff; padding: 12px 25px; border-radius: 30px; font-size: 18px; font-weight: bold; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
                    AI評価スコア： ${scoreText ? scoreText : '-'}
                  </span>
                </div>
              </div>
              <div style="background-color: #f1f3f4; padding: 15px; text-align: center; font-size: 12px; color: #666;">
                ※このメールはシステムからの自動送信です。<br>質問がある場合は、次回の授業で先生に聞いてみましょう！
              </div>
            </div>
          `;

          try {
            GmailApp.sendEmail(emailAddress, subject, plainBody, { htmlBody: htmlBody });
            sendCount++;
            const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd HH:mm');
            sheet.getRange(i + Config.DATA_START_ROW, Config.COL_MAIL_STATUS).setValue(`${dateStr} 送信済`);
          } catch(e) {
            sheet.getRange(i + Config.DATA_START_ROW, Config.COL_MAIL_STATUS).setValue(`送信エラー`);
          }
        }
      }
    }
    
    // ★ 追加: 出っぱなしのトーストを「完了（5秒で自動消去）」で上書きして消す
    ss.toast('メールの送信処理が完了しました。', '✨ 完了', 5);
    ui.alert(`送信完了\n\n新しく ${sendCount} 名の生徒にフィードバックメールを送信しました！`);
  },

  sendReminders: function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(Config.MAIN_SHEET);
    const ui = SpreadsheetApp.getUi();
    const rosterSheet = ss.getSheetByName(Config.ROSTER_SHEET);

    if (!sheet) { ui.alert(`エラー: 「${Config.MAIN_SHEET}」シートが見つかりません。`); return; }
    if (!rosterSheet) { ui.alert('エラー: 名簿シートが見つかりません。'); return; }

    const response = ui.alert('確認', `現在「${Config.MAIN_SHEET}」シート上で英作文が空欄になっている生徒全員に対して、提出を促すリマインドメールを送信しますか？`, ui.ButtonSet.YES_NO);
    if (response !== ui.Button.YES) return;

    ss.toast('リマインドメールを送信しています...', '🔔 送信中', -1);

    const rosterData = rosterSheet.getDataRange().getValues();
    const emailMap = {};
    for (let i = 1; i < rosterData.length; i++) {
      if (rosterData[i][0] && rosterData[i][1]) emailMap[rosterData[i][0]] = rosterData[i][1];
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < Config.DATA_START_ROW) return;
    const data = sheet.getRange(Config.DATA_START_ROW, 1, lastRow - Config.DATA_START_ROW + 1, Config.COL_STUDENT).getValues();
    let sendCount = 0;

    for (let i = 0; i < data.length; i++) {
      const num = data[i][0];
      const studentText = data[i][Config.COL_STUDENT - 1];

      if (num && !studentText) {
        const emailAddress = emailMap[num];
        if (emailAddress) {
          const subject = "【英語科より】英作文提出のお願い";
          const body = `お疲れ様です。英語科からの連絡です。\n\n現在、あなたの英作文がまだ提出されていません。\nGoogleフォームから、早めに提出をお願いします。\n\n※すでに提出済みの場合は、行き違いですのでご容赦ください。\n※このメールはシステムからの自動送信です。`;
          try { GmailApp.sendEmail(emailAddress, subject, body); sendCount++; } catch(e) {}
        }
      }
    }
    
    // ★ 追加: 出っぱなしのトーストを「完了（5秒で自動消去）」で上書きして消す
    ss.toast('リマインド送信処理が完了しました。', '✨ 完了', 5);
    ui.alert(`送信完了\n\n未提出の生徒 ${sendCount} 名にリマインドメールを送信しました！`);
  }
};