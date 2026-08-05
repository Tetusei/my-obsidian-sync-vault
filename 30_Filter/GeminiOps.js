/**
 * @file GeminiOps.gs
 * @description Gemini API(gemini-flash-latest)と安全に通信し、検知メールの高度な文脈解析と緊急度判定を行います。
 */

/**
 * Gemini APIを呼び出して、検知内容を自動分析する
 * @param {string} subject メールの件名
 * @param {string} reason 検知理由（検索キーワードなど）
 * @param {string} snippet メールの本文抜粋（文脈）
 * @return {Object} 分析結果オブジェクト { urgency, intention, summary }
 */
function analyzeDetectionWithGemini(subject, reason, snippet) {
  try {
    const apiKey = getStoredApiKey();
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${Config.GEMINI_MODEL}:generateContent?key=${apiKey}`;
    
    // AIへの綿密な指示（学校現場用のプロンプト）
    const prompt = `
あなたは学校のICT安全管理アシスタントです。
生徒の端末利用において、Webフィルターで検知された検索履歴やアクセスログの通知メールを分析し、教員が直感的に状況を把握できるようにしてください。

【検知データ】
件名: ${subject}
検知理由（キーワード等）: ${reason}
検知された文脈・周辺テキスト: ${snippet}

【指示】
以下の3つの項目を解析し、指定のJSONフォーマットのみで出力してください。解説や余計な文章は絶対に含めないでください。

1. urgency (緊急度): 「高」「中」「低」のいずれか1文字。
   - 「高」: 自傷行為・自殺予兆、暴力、重大ないじめ、犯罪・非行に直結する危険な内容、性的過激表現
   - 「中」: 授業に関係のないエンタメ、ゲーム、アニメ、軽微な規律違反、悩みの兆候、アダルトライト層
   - 「低」: 誤検知の可能性が高いもの、一般的な言葉の調べ学習、通常のニュース閲覧の範囲
2. intention (推測される意図): 生徒がどのような目的や心理でこれを見た・検索したかの短い推察（30文字以内）
3. summary (先生への要約): 何が起きたかを教員向けに分かりやすく解説した一言要約（50文字以内）

【出力フォーマット（厳守）】
{
  "urgency": "緊急度を配置",
  "intention": "推測される意図を配置",
  "summary": "先生への要約を配置"
}
`;

    const payload = {
      contents: [{
        parts: [{ text: prompt }]
      }],
      generationConfig: {
        responseMimeType: "application/json" // 返答をJSONに固定
      }
    };

    const options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode !== 200) {
      console.error(`Gemini API エラー (コード: ${responseCode}): ` + response.getContentText());
      return { urgency: '不明', intention: 'AI通信エラー', summary: 'AI解析の呼び出しに失敗しました。' };
    }

    const resJson = JSON.parse(response.getContentText());
    const resultText = resJson.candidates[0].content.parts[0].text;
    
    // AIの返答を解析してオブジェクトとして返す
    const parsedResult = JSON.parse(resultText);
    return {
      urgency: parsedResult.urgency || '中',
      intention: parsedResult.intention || '判定不可',
      summary: parsedResult.summary || '詳細要約の取得に失敗しました。'
    };

  } catch (e) {
    console.error('Gemini解析処理内で例外が発生しました: ' + e.message);
    return { urgency: '不明', intention: '解析エラー', summary: '一時的なエラーにより解析できませんでした。' };
  }
}