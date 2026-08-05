/**
 * @file Main.gs
 * @description メールの自動巡回、AI分析依頼、シート記録、チャット通知の全工程をコントロールします。
 */

function checkGmailAndNotify() {
  let webhookUrl;
  try {
    webhookUrl = getWebhookUrl();
    if (!webhookUrl) return;
  } catch (e) {
    console.error("Webhookの取得に失敗: " + e.message);
    return;
  }

  // 基礎データシートのB3セルから対象者を取得
  let targetPerson = "";
  try {
    targetPerson = getTargetPerson();
  } catch (e) {
    console.error("対象者フィルターの取得に失敗: " + e.message);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // 画面の右下にトースト通知を出し、先生に稼働中であることを示して安心感を提供
  ss.toast("見守りフィルターの未読メールをスキャン中...", "システム巡回中");

  const studentDict = getStudentDictionary();
  const threads = GmailApp.search(Config.SEARCH_QUERY);

  if (threads.length > 0) {
    let processedCount = 0;

    for (const thread of threads) {
      const messages = thread.getMessages();
      for (const message of messages) {
        
        if (message.isUnread()) {
          const subject = message.getSubject();
          
          // 件名に対象者フィルターの文字列が含まれているか判定
          if (targetPerson && !subject.includes(targetPerson)) {
            continue; // 対象者が含まれない場合は処理をスキップ（未読のまま維持）
          }

          const body = message.getPlainBody();
          const mailDate = message.getDate();
          const formattedDate = Utilities.formatDate(mailDate, "JST", "yyyy/MM/dd HH:mm");
          
          // 1. 件名からメールアドレスを抽出して生徒情報を特定
          const emailMatch = subject.match(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/);
          let studentDisplay = "⚠️ 名簿に該当なし";
          let emailStr = "取得失敗";
          let studentInfo = null;

          if (emailMatch) {
            emailStr = emailMatch[0].toString().trim();
            studentInfo = studentDict[emailStr];
            if (studentInfo) {
               studentDisplay = `🎓 ${studentInfo.grade}年 ${studentInfo.class}組 ${studentInfo.number}番 ${studentInfo.name}`;
            } else {
               studentDisplay = `❓ 名簿未登録 (${emailStr})`;
            }
          }

          // 2. 本文から「検知理由」と「アクセス回数」を正規表現で正確に抽出
          const reasonMatch = body.match(/メール送信対象理由：(.*?)(?:\n|\r|$)/);
          const reason = reasonMatch ? reasonMatch[1].trim() : "取得失敗";

          const countMatch = body.match(/アクセス回数：(.*?)(?:\n|\r|$)/);
          const accessCount = countMatch ? countMatch[1].trim() : "取得失敗";

          // 本文の最初の150文字を抜粋（AIの分析材料およびチャット表示用）
          const snippet = body.substring(0, 150).replace(/\n/g, " ");

          // 3. Gemini API（固定モデル: gemini-flash-latest）を呼び出して高度なAI判定を実行
          const aiResult = analyzeDetectionWithGemini(subject, reason, snippet);

          // 4. 「R8年度」シートの最終行へ自動記録を実行
          recordToLogSheet(mailDate, studentInfo, emailStr, reason, accessCount, aiResult);

          // 5. AIが判定した「緊急度」に基づいて、チャット通知の先頭のマークを自動選定
          let urgencyBadge = "ℹ️【情報】";
          if (aiResult.urgency === "高") {
            urgencyBadge = "🚨🚨🚨【至急確認・最優先対応】";
          } else if (aiResult.urgency === "中") {
            urgencyBadge = "⚠️【要注意確認】";
          }

          // 6. Google Chat送信用メッセージ（高度判定版）の組み立て
          const payload = {
            'text': `*${urgencyBadge} 見守りフィルター・AI高度解析速報*
> *対象生徒:* ${studentDisplay}
> *検知理由:* ${reason} (アクセス: ${accessCount}回)
> 
> *🤖 人工知能（Gemini）による分析結果*
> ・*AI緊急度:* [ *${aiResult.urgency}* ]
> ・*推測される意図:* ${aiResult.intention}
> • *先生への一言要約:* ${aiResult.summary}
> 
> *【受信メール詳細】*
> ・*件名:* ${subject}
> ・*受信日時:* ${formattedDate}
> ・*本文の抜粋:* ${snippet}...
> 
> *📬 直接このメールを開いて確認する:* > https://mail.google.com/mail/u/0/#inbox/${message.getId()}`
          };

          const options = {
            'method': 'post',
            'contentType': 'application/json; charset=UTF-8',
            'payload': JSON.stringify(payload)
          };

          // Google Chat (Webhook URL) へ一斉送信
          UrlFetchApp.fetch(webhookUrl, options);
          
          // 既読に設定して次回巡回時の重複通知・重複記録を鉄壁にガード
          message.markRead(); 
          processedCount++;
        }
      }
    }
    
    if (processedCount > 0) {
      ss.toast(`${processedCount} 件の新しい検知メールを処理し、R8年度シートへの記録とチャット通知を完了しました。`, "✅ 処理完了");
    }
  } else {
    // 未読がゼロだった場合もトーストで静かに着信なしをお知らせ
    ss.toast("現在、新しく検知された未読のメールはありません。", "💤 巡回終了", 3);
  }
}