/**
 * Main.gs
 * Description: AI所見作成、推敲、日々の記録からの全自動生成機能（ダッシュボード連携版）
 * Version: v2.5.0 (AI処理のエラー耐性強化・全行一括処理を追加)
 */

function generateFromDailyRecords() {
  generateFromDailyRecordsInternal(false);
}

function generateFromDailyRecordsAll() {
  generateFromDailyRecordsInternal(true);
}

function generateFromDailyRecordsInternal(allRows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const sheetName = sheet.getName();
  const ui = SpreadsheetApp.getUi();

  if (!CONFIG.SHEET_TARGETS.includes(sheetName)) {
    ui.alert(`「1学期〜3学期」シートを開いた状態で実行してください。`);
    return;
  }

  let startRow, lastRow;
  if (allRows) {
    startRow = 2;
    lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      ui.alert('生徒データが見つかりません。');
      return;
    }
  } else {
    const activeRange = sheet.getActiveRange();
    startRow = activeRange.getRow();
    const numRows = activeRange.getNumRows();
    lastRow = startRow + numRows - 1;

    if (startRow === 1) return ui.alert('見出し行が選択されています。データのある行を選択してください。');
  }

  const baseSheet = ss.getSheetByName(CONFIG.SHEET_BASE);
  let startCell, endCell;
  if (sheetName === '1学期') { startCell = CONFIG.CELL_TERM1_START; endCell = CONFIG.CELL_TERM1_END; }
  else if (sheetName === '2学期') { startCell = CONFIG.CELL_TERM2_START; endCell = CONFIG.CELL_TERM2_END; }
  else if (sheetName === '3学期') { startCell = CONFIG.CELL_TERM3_START; endCell = CONFIG.CELL_TERM3_END; }

  const termStart = new Date(baseSheet.getRange(startCell).getValue());
  const termEnd = new Date(baseSheet.getRange(endCell).getValue());
  termEnd.setHours(23, 59, 59, 999);

  if (isNaN(termStart.getTime()) || isNaN(termEnd.getTime())) {
    return ui.alert('エラー', `基礎データシートの期間設定が正しくありません。\n「YYYY/MM/DD」形式で入力してください。`, ui.ButtonSet.OK);
  }

  const targetMsg = allRows ? "全生徒" : `${startRow}行目 から ${lastRow}行目 の対象生徒`;
  if (ui.alert('確認', `【${sheetName}】\n${targetMsg} について、日々の記録から自動生成を行います。\nよろしいですか？`, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  ss.toast(`処理を開始します。`, '開始', 5);
  const startTime = Date.now();

  try {
    const schoolType = String(baseSheet.getRange(CONFIG.CELL_SCHOOL_TYPE).getValue() || "");
    const charLimit = baseSheet.getRange(CONFIG.CELL_CHAR_LIMIT).getValue();
    const additionalPrompt = baseSheet.getRange(CONFIG.CELL_PROMPT).getValue();
    const ngWords = baseSheet.getRange(CONFIG.CELL_NG_WORDS).getValue();
    const localRules = baseSheet.getRange(CONFIG.CELL_RULES).getValue();
    
    let rulePrompt = (ngWords || localRules) ? "\n【学校独自のルール】\n" + (localRules ? `・ルール: ${localRules}\n` : "") + (ngWords ? `・NGワード: ${ngWords}\n` : "") : "";
    let toneInstruction = schoolType.includes("中学校") ? "中学校の教員として客観的な表現" : "小学校の教員として温かみのある平易な表現";

    // 名簿データを一括取得（固定情報の読み取り用）
    const meiboSheet = ss.getSheetByName(CONFIG.SHEET_MEIBO);
    const meiboData = meiboSheet.getDataRange().getValues();

    // 成績シートの一括取得（存在する場合）
    const seisekiSheet = ss.getSheetByName(CONFIG.SHEET_SEISEKI || '成績');
    let seisekiData = null;
    if (seisekiSheet && seisekiSheet.getLastRow() >= 3) {
      seisekiData = seisekiSheet.getDataRange().getValues();
    }

    for (let r = startRow; r <= lastRow; r++) {
      if (Date.now() - startTime > CONFIG.EXECUTION_TIME_LIMIT) break;

      const number = sheet.getRange(r, 1).getValue();
      const name = sheet.getRange(r, 2).getValue();
      if (!number || !name) continue;

      const studentSheet = ss.getSheetByName(`${number}_${name}`);
      const aiCell = sheet.getRange(r, CONFIG.COL_AI_COMMENT);
      
      if (!studentSheet) {
        aiCell.setValue("※個人シートがありません");
        continue;
      }

      // 🌟 名簿から委員会・クラブ・特記事項を取得
      let committee = "";
      let club = "";
      let specialNote = "";
      for (let m = 1; m < meiboData.length; m++) {
        if (meiboData[m][0] === number && meiboData[m][1] === name) {
          committee = meiboData[m][3] || ""; // D列: 委員会
          club = meiboData[m][4] || "";      // E列: クラブ
          specialNote = meiboData[m][5] || "";// F列: 特記事項
          break;
        }
      }

      // 🌟 成績シートからM列（まとめ出力）を取得
      let seisekiSummaryText = "";
      if (seisekiData) {
        for (let s = 2; s < seisekiData.length; s++) {
          if (seisekiData[s][0] === number && seisekiData[s][1] === name) {
            seisekiSummaryText = seisekiData[s][12] || ""; // M列 (13列目)
            break;
          }
        }
      }

      let fixedInfoText = "";
      if (committee || club || specialNote || seisekiSummaryText) {
        fixedInfoText = `\n【生徒の基本情報・学力テスト結果（特徴の参考として考慮に含めてください）】\n`;
        if (committee) fixedInfoText += `・所属委員会: ${committee}\n`;
        if (club) fixedInfoText += `・所属クラブ・部活動: ${club}\n`;
        if (specialNote) fixedInfoText += `・特記事項（重要）: ${specialNote}\n`;
        if (seisekiSummaryText) fixedInfoText += `・標準学力テスト成績:\n${seisekiSummaryText}\n`;
      }

      const lastRecRow = studentSheet.getLastRow();
      const lastRecCol = studentSheet.getLastColumn();
      if (lastRecRow > 1) {
        studentSheet.getRange(2, 1, lastRecRow - 1, lastRecCol).sort({column: 1, ascending: true});
      }

      const recordsData = studentSheet.getDataRange().getValues();
      let compiledRecords = [];
      
      for (let i = 1; i < recordsData.length; i++) {
        const rDate = new Date(recordsData[i][0]);
        const recordText = recordsData[i][1];
        const priorityMark = recordsData[i][2]; 
        
        if (rDate >= termStart && rDate <= termEnd && recordText) {
          let recordLine = `・[${Utilities.formatDate(rDate, "JST", "M/d")}] ${recordText}`;
          if (priorityMark === '〇' || priorityMark === '○' || priorityMark === '◯') {
            recordLine = `【★最優先】 ` + recordLine;
          }
          compiledRecords.push(recordLine);
        }
      }

      if (compiledRecords.length === 0 && fixedInfoText === "") {
        aiCell.setValue("※指定期間内の記録や基本情報がありません");
        // 【修正追加】記録がない場合は、古い特徴量（5列分）もクリアする
        sheet.getRange(r, CONFIG.COL_FEATURES_START, 1, 5).clearContent();
        continue;
      }

      const prompt = `あなたは優秀な教員です。以下の「日々の記録」および「生徒の基本情報」から、実際に記載されている事実や言葉のみを抽出し、通知表に記載すべき特徴を最大5つ挙げてください。また、それをもとに所見の文章を作成してください。

【厳守事項（特徴の抽出について）】
・「【★最優先】」というマークがついている記録は、先生が必ず通知表に載せたいと指定した最重要項目です。必ず特徴（feature1〜5のいずれか）として抽出してください。
・先生が入力した「具体的な言葉」や事実は、勝手に意訳せず、極力そのまま特徴として抽出すること。
・記録の件数や情報量が少ない場合、無理に5つの特徴を作らないこと。残りの特徴枠は絶対に空欄（""）にすること。想像で捏造しないこと。
・記録が多数ある場合は、【★最優先】のものを中心に、通知表にふさわしい重要なものを最大5つ厳選すること。

【厳守事項（所見の作成について）】
・${toneInstruction}
・所見の全体の文字数を【 ${charLimit - 5}文字以上、${charLimit + 5}文字以内】に調整することを目標とする。
・ただし、記録の情報量が少なく指定文字数に達することが不可能な場合は、文字数制限を無視して、抽出した事実のみで簡潔な短い文章を作成すること。想像で話を膨らませないこと。
・個人名（「〜さん」など）を絶対に含めないこと。
・追加の指示: ${additionalPrompt || '特になし'}
${rulePrompt}
${fixedInfoText}
【対象生徒の日々の記録】
${compiledRecords.join('\n')}

【出力形式】
必ず以下のJSONスキーマに従い、JSON形式でのみ出力してください。
{
  "feature1": "抽出した特徴1（記録から事実のみを記載）",
  "feature2": "抽出した特徴2（ない場合は必ず空欄にする）",
  "feature3": "抽出した特徴3（ない場合は必ず空欄にする）",
  "feature4": "抽出した特徴4（ない場合は必ず空欄にする）",
  "feature5": "抽出した特徴5（ない場合は必ず空欄にする）",
  "comment": "作成した所見の文章（事実のみで構成）"
}`;

      aiCell.setValue(`⏳ 【${name}】さんの記録をAIが読み込み中...`).setBackground('#fff2cc');
      SpreadsheetApp.flush();

      try {
        const aiResponse = callGeminiAPI(prompt, baseSheet, true);
        // 【修正追加】JSON.parse を Utils.gs の safeParseJSON に変更し、エラーを回避
        const resultJson = safeParseJSON(aiResponse);
        
        sheet.getRange(r, CONFIG.COL_FEATURES_START, 1, 5).setVerticalAlignment('middle');
        sheet.getRange(r, CONFIG.COL_FEATURES_START).setValue(resultJson.feature1 || "");
        sheet.getRange(r, CONFIG.COL_FEATURES_START + 1).setValue(resultJson.feature2 || "");
        sheet.getRange(r, CONFIG.COL_FEATURES_START + 2).setValue(resultJson.feature3 || "");
        sheet.getRange(r, CONFIG.COL_FEATURES_START + 3).setValue(resultJson.feature4 || "");
        sheet.getRange(r, CONFIG.COL_FEATURES_START + 4).setValue(resultJson.feature5 || "");
        
        aiCell.setValue(resultJson.comment || "").setBackground(null).setVerticalAlignment('middle');
      } catch (e) {
        aiCell.setValue(`⚠️ AI処理エラー: ${e.message}`).setBackground('#f8cccc');
        SpreadsheetApp.flush();
      }
    }
    ss.toast(`完了しました！`, '完了', 5);
  } catch (error) {
    ui.alert('全体エラー:\n' + error.message);
  }
}

function generateAIComments() {
  generateAICommentsInternal(false);
}

function generateAICommentsAll() {
  generateAICommentsInternal(true);
}

function generateAICommentsInternal(allRows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const sheetName = sheet.getName();
  const ui = SpreadsheetApp.getUi();

  const isReportCard = CONFIG.SHEET_TARGETS.includes(sheetName);
  const isYouroku = (sheetName === CONFIG.SHEET_YOUROKU);
  const isAnnual = (sheetName === CONFIG.SHEET_ANNUAL);

  if (!isReportCard && !isYouroku && !isAnnual) return;

  let startRow, lastRow;
  if (allRows) {
    startRow = 2;
    lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
  } else {
    const activeRange = sheet.getActiveRange();
    startRow = activeRange.getRow();
    const numRows = activeRange.getNumRows();
    lastRow = startRow + numRows - 1;
    if (startRow === 1) return;
  }

  // セキュリティチェック（個人名の混入防止）
  const warningRows = scanForPersonalInfo(sheet, startRow, lastRow, isReportCard, isYouroku || isAnnual, CONFIG.COL_FEATURES_START, CONFIG.COL_FEATURES_END);
  if (warningRows.length > 0) {
    ui.alert(
      '🚨 セキュリティブロック', 
      `以下の行の特徴入力に「〜さん」「〜君」などの個人名が含まれている可能性があります。\n\n【該当行】: ${warningRows.join(', ')}行目\n\n情報漏洩防止のため、処理を中断しました。\n匿名表現（「Aさん」「生徒」など）に修正してから再度実行してください。`, 
      ui.ButtonSet.OK
    );
    return; 
  }

  const targetMsg = allRows ? "全生徒" : "選択行の生徒";
  if (ui.alert('確認', `${targetMsg}の所見を作成しますか？`, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  const startTime = Date.now();

  try {
    const baseSheet = ss.getSheetByName(CONFIG.SHEET_BASE);
    // 【修正追加】学校種別が空欄でもエラーにならないよう String() で囲む
    const schoolType = String(baseSheet.getRange(CONFIG.CELL_SCHOOL_TYPE).getValue() || "");
    const ngWords = baseSheet.getRange(CONFIG.CELL_NG_WORDS).getValue();
    const localRules = baseSheet.getRange(CONFIG.CELL_RULES).getValue();
    let rulePrompt = (ngWords || localRules) ? "\n【表記ルール】\n" + (localRules ? `・ルール: ${localRules}\n` : "") + (ngWords ? `・NGワード: ${ngWords}\n` : "") : "";

    for (let r = startRow; r <= lastRow; r++) {
      if (Date.now() - startTime > CONFIG.EXECUTION_TIME_LIMIT) break;
      
      let prompt = "";
      let targetCol = 0;

      // ----------------------------------------------------
      // 通知表の所見作成（1学期〜3学期シート）
      // ----------------------------------------------------
      if (isReportCard) {
        const charLimit = baseSheet.getRange(CONFIG.CELL_CHAR_LIMIT).getValue();
        let features = [];
        for (let c = CONFIG.COL_FEATURES_START; c <= CONFIG.COL_FEATURES_END; c++) {
          const val = sheet.getRange(r, c).getValue();
          if (val) features.push(val);
        }
        if (features.length === 0) continue;

        prompt = `${schoolType.includes("中学校") ? "中学校の教員として客観的な表現" : "小学校の教員として温かみのある平易な表現"}
【条件】全体の文字数を【必ず ${charLimit - 5}文字以上、${charLimit + 5}文字以内】に調整。
${rulePrompt}
【対象者の特徴】\n${features.join('、 ')}\n出力は所見の文章のみとしてください。`;
        targetCol = CONFIG.COL_AI_COMMENT;

      // ----------------------------------------------------
      // 指導要録の作成（要録シート）
      // ----------------------------------------------------
      } else if (isYouroku) {
        const term1 = sheet.getRange(r, CONFIG.COL_Y_TERM1).getValue();
        const term2 = sheet.getRange(r, CONFIG.COL_Y_TERM2).getValue();
        const term3 = sheet.getRange(r, CONFIG.COL_Y_TERM3).getValue();
        if (!term1 && !term2 && !term3) continue;

        const charLimit = baseSheet.getRange(CONFIG.CELL_YOUROKU_CHAR_LIMIT).getValue();
        const additionalPrompt = baseSheet.getRange(CONFIG.CELL_YOUROKU_PROMPT).getValue();
        
        let yourokuFormatPrompt = schoolType.includes("中学校") ? CONFIG.PROMPT_YOUROKU_JHS : CONFIG.PROMPT_YOUROKU_ELEM;

        prompt = `${yourokuFormatPrompt}
【条件】全体の文字数を【必ず ${charLimit - 5}文字以上、${charLimit + 5}文字以内】に調整すること。
${rulePrompt}
${additionalPrompt ? `【追加の指示】\n${additionalPrompt}\n` : ""}
【通知表データ（この内容を要約・再構成してください）】
1学期: ${term1}
2学期: ${term2}
3学期: ${term3}

出力は作成した指導要録の文章のみとしてください。`;
        targetCol = CONFIG.COL_Y_AI;

      // ----------------------------------------------------
      // 年間通知表所見の作成（年間シート）
      // ----------------------------------------------------
      } else if (isAnnual) {
        const term1 = sheet.getRange(r, CONFIG.COL_Y_TERM1).getValue();
        const term2 = sheet.getRange(r, CONFIG.COL_Y_TERM2).getValue();
        const term3 = sheet.getRange(r, CONFIG.COL_Y_TERM3).getValue();
        if (!term1 && !term2 && !term3) continue;

        const charLimit = baseSheet.getRange(CONFIG.CELL_YOUROKU_CHAR_LIMIT).getValue() || baseSheet.getRange(CONFIG.CELL_CHAR_LIMIT).getValue();
        const additionalPrompt = baseSheet.getRange(CONFIG.CELL_YOUROKU_PROMPT).getValue();
        
        let toneInstruction = schoolType.includes("中学校") ? 
          "中学校の教員として、保護者に向けて生徒の1年間の学習成果や成長、活躍の様子を具体的かつ丁寧に伝える表現" : 
          "小学校の教員として、保護者に向けて児童の1年間の頑張りや成長を温かく称え、前向きに伝える表現";

        prompt = `あなたは教員です。年に1回保護者へ発行する「年間通知表の所見」を作成してください。
1学期・2学期・3学期の記録や様子をもとに、児童・生徒の1年間の学習成果、学校生活・特別活動での活躍、行動面での成長を総合的に評価し、保護者へ向けて温かく前向きに伝える通知表の文章を作成してください。

【執筆ルール（年間通知表スタイル）】
・${toneInstruction}とすること。
・文末表現は通知表に相応しい標準的な丁寧表現（「〜することができました。」「〜に努める姿が見られました。」「〜成長が感じられました。」等）を用いること。
・指導要録のような硬い行政記録表現（「〜を遂行した」等）や、箇条書き・中黒（・）は使わず、1つの自然につなげた文章に仕上げること。
・全体の文字数を【必ず ${charLimit - 5}文字以上、${charLimit + 5}文字以内】に調整すること。
・個人名（「〜さん」「〜君」など）を絶対に含めないこと。
${rulePrompt}
${additionalPrompt ? `【追加の指示】\n${additionalPrompt}\n` : ""}
【各学期の記録・メモデータ】
1学期: ${term1}
2学期: ${term2}
3学期: ${term3}

出力は作成した年間通知表所見の文章のみとしてください。`;
        targetCol = CONFIG.COL_Y_AI;
      }

      const targetCell = sheet.getRange(r, targetCol);
      targetCell.setValue(`⏳ AI処理中...`).setBackground('#fff2cc');
      SpreadsheetApp.flush();

      try {
        const aiResponse = callGeminiAPI(prompt, baseSheet, false);
        targetCell.setValue(aiResponse).setBackground(null).setVerticalAlignment('middle');
      } catch (e) {
        targetCell.setValue(`⚠️ AI処理エラー: ${e.message}`).setBackground('#f8cccc');
      }
    }
    ss.toast(`完了しました！`, '完了', 5);
  } catch (error) {
    ui.alert('全体エラー:\n' + error.message);
  }
}

function proofreadComments() {
  proofreadCommentsInternal(false);
}

function proofreadCommentsAll() {
  proofreadCommentsInternal(true);
}

function proofreadCommentsInternal(allRows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const sheetName = sheet.getName();
  const ui = SpreadsheetApp.getUi();

  const isReportCard = CONFIG.SHEET_TARGETS.includes(sheetName);
  const isYouroku = (sheetName === CONFIG.SHEET_YOUROKU);
  const isAnnual = (sheetName === CONFIG.SHEET_ANNUAL);
  if (!isReportCard && !isYouroku && !isAnnual) return;

  let startRow, lastRow;
  if (allRows) {
    startRow = 2;
    lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
  } else {
    const activeRange = sheet.getActiveRange();
    startRow = activeRange.getRow();
    const numRows = activeRange.getNumRows();
    lastRow = startRow + numRows - 1;
    if (startRow === 1) return;
  }

  const targetMsg = allRows ? "全生徒" : "選択行の生徒";
  if (ui.alert('確認', `${targetMsg}の推敲を実施しますか？`, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  const startTime = Date.now();

  try {
    const baseSheet = ss.getSheetByName(CONFIG.SHEET_BASE);
    for (let r = startRow; r <= lastRow; r++) {
      if (Date.now() - startTime > CONFIG.EXECUTION_TIME_LIMIT) break;

      let finalComment = "";
      let targetCol = 0;
      if (isReportCard) {
        finalComment = sheet.getRange(r, CONFIG.COL_FINAL_COMMENT).getValue();
        targetCol = CONFIG.COL_PROOFREAD;
      } else if (isYouroku || isAnnual) {
        finalComment = sheet.getRange(r, CONFIG.COL_Y_FINAL).getValue();
        targetCol = CONFIG.COL_Y_PROOFREAD;
      }
      
      if (!finalComment) continue;

      const prompt = `あなたは優秀な校正者です。以下の文章を添削してください。
【重要】修正の理由や挨拶は出力しない。修正後の文章のみを出力する。変更箇所のみを <red> と </red> で囲む。変更がない場合はタグを使わずそのまま出力する。
【対象の文章】\n${finalComment}`;

      const targetCell = sheet.getRange(r, targetCol);
      targetCell.setValue(`⏳ 推敲中...`).setBackground('#fff2cc');
      SpreadsheetApp.flush();

      try {
        const aiResponse = callGeminiAPI(prompt, baseSheet, false);
        const plainText = aiResponse.replace(/<\/?red>/g, "");
        targetCell.setBackground(null);
        
        if (plainText === finalComment) {
          targetCell.setValue(finalComment).setVerticalAlignment('middle');
          continue; 
        }
        
        const builder = SpreadsheetApp.newRichTextValue().setText(plainText);
        const redStyle = SpreadsheetApp.newTextStyle().setForegroundColor("red").build();
        const regex = /<red>(.*?)<\/red>/g;
        let match; let offset = 0; 
        
        while ((match = regex.exec(aiResponse)) !== null) {
          const start = match.index - offset;
          const end = start + match[1].length;
          builder.setTextStyle(start, end, redStyle);
          offset += 11; 
        }
        targetCell.setRichTextValue(builder.build()).setVerticalAlignment('middle');
      } catch (e) {
        targetCell.setValue(`⚠️ AI処理エラー: ${e.message}`).setBackground('#f8cccc');
      }
    }
    ss.toast('完了しました！', '完了', 5);
  } catch (error) {
    ui.alert('全体エラー:\n' + error.message);
  }
}