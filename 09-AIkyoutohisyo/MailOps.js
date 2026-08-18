/**
 * MailOps.gs
 * メール自動解析（安全対策・レートリミット・エラー回避版）
 */

function fetchAndFilterMail() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const todoSheet = ss.getSheetByName(Config.SHEET_NAME_TODO);
  const masterSheet = ss.getSheetByName(Config.SHEET_NAME_MASTER);
  
  const apiKey = getApiKey();
  const targetEmail = masterSheet.getRange(Config.MASTER_POS.TARGET_EMAIL_CELL).getValue();
  const modelName = masterSheet.getRange(Config.MASTER_POS.MODEL_NAME_CELL).getValue();

  console.log(`[設定確認] APIキー: ${apiKey ? "〇取得" : "❌空欄"}, モデル: ${modelName}`);
  if (!apiKey || !modelName) {
    console.error("APIキーまたはモデル名が設定されていません。処理を中断します。");
    return;
  }

  let searchQuery = 'is:unread';
  if (targetEmail) {
    searchQuery = `deliveredto:${targetEmail} is:unread`;
  }
  console.log(`[検索クエリ] ${searchQuery}`);
  
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) {
    console.log('メール解析は既に実行中のため、今回の処理をスキップします。');
    return;
  }

  try {
    const threads = GmailApp.search(searchQuery, 0, Config.MAIL_FETCH_LIMIT);
    console.log(`[発見した未読スレッド数] ${threads.length} 通`);
    if (threads.length === 0) return; 

    // メルアドシートがなければ自動作成
    ensureMailAddressSheet();

    // threads.forEach から forループ に変更（途中のスキップや休憩を制御するため）
    for (let i = 0; i < threads.length; i++) {
      const thread = threads[i];
      // スレッド先頭ではなく、最新の未読メッセージを1件ずつ処理する
      const message = thread.getMessages().reverse().find(item => item.isUnread());
      if (!message) continue;
      const subject = message.getSubject();
      const body = message.getPlainBody();
      const from = message.getFrom();
      const hasAttachments = message.getAttachments().length > 0;
      const receptionDate = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm");

      console.log(`[解析開始 (${i + 1}/${threads.length})] 件名: ${subject}`);

      // ✨ エラーメールやGAS関連通知メールのスキップ処理
      const fromLower = from.toLowerCase();
      const subjectLower = subject.toLowerCase();
      
      // システムエラーメールの判定
      const isErrorMail = 
        fromLower.includes('mailer-daemon') ||
        fromLower.includes('postmaster') ||
        subjectLower.includes('delivery status notification') ||
        subjectLower.includes('undelivered mail') ||
        subjectLower.includes('failure notice') ||
        subjectLower.includes('mail delivery failed');
        
      // GAS関連・システム通知メールの判定
      const isGasMail =
        subjectLower.includes('apps script') ||
        subjectLower.includes('apps-script') ||
        fromLower.includes('apps-script-notifications') ||
        subjectLower.includes('script.google.com') ||
        subjectLower.includes('google apps script');
      
      if (isErrorMail || isGasMail) {
        const reason = isErrorMail ? "配信不能通知等のシステムエラーメール" : "Google Apps Script 関連のシステム通知メール";
        console.log(`⚪ [システムメール除外] 「${subject}」は${reason}のため、既読にしてスキップします。`);
        message.markRead();
        continue;
      }

      // 同時に届いたメールのすり抜け（GASの書き込み遅延）を防ぐため強制同期
      SpreadsheetApp.flush();

      // 1. 同一日の重複メール判定（異なるルートから転送された同一メール of スキップ）
      if (isDuplicateEmail(todoSheet, subject, body, receptionDate)) {
        console.log(`⚪ [重複検知] 「${subject}」は同日内に既に登録されているため、登録をスキップし既読にします。`);
        message.markRead();
        continue;
      }

      const analysis = askGemini(subject, body, modelName);

      // 🚨 安全装置：AIからの応答が正常に取得できなかった場合は、未読のまま残して次のメールへ
      if (analysis === null) {
        console.error(`❌ [解析失敗] Gemini APIから正常な応答が得られなかったため、既読にせずスキップします。次の実行時に再トライします。`);
        continue; 
      }

      const mailLink = `https://mail.google.com/mail/u/0/#inbox/${message.getId()}`;

      let finalDueDate = "";
      if (analysis.dueDate) {
        const datePattern = /^\d{4}\/\d{1,2}\/\d{1,2}$/;
        if (datePattern.test(analysis.dueDate)) {
          finalDueDate = analysis.dueDate;
        }
      }

      // 添付ファイルがある場合はタイトルの先頭に 📎 アイコンを付与
      const taskTitle = (hasAttachments ? '📎 ' : '') + (analysis.title || subject);
      const taskContent = analysis.summary || body.substring(0, 500);
      const taskPic = analysis.pic || '教頭';
      
      // 全メールを取り込む。タスク外のメールはデフォルトで「手元で留める（転送不要）」に設定
      const isTask = analysis.isTask;
      const stakeholder = isTask ? (analysis.stakeholder || '手元で留める（転送不要）') : '手元で留める（転送不要）';
      const correspondingName = getNameByRole(stakeholder) || '';

      // 件名をハイパーリンク式にする
      const titleFormula = `=HYPERLINK("${mailLink}", "${taskTitle.replace(/"/g, '""')}")`;

      // 📎 添付ファイルがある場合はフォルダに保存してリンクを取得
      let attachmentText = "";
      if (hasAttachments) {
        try {
          const attachmentsFolderUrl = masterSheet.getRange(Config.MASTER_POS.ATTACHMENTS_FOLDER_URL_CELL).getValue();
          let attachmentsFolder;
          if (attachmentsFolderUrl && attachmentsFolderUrl.includes('drive.google.com/')) {
            attachmentsFolder = DriveApp.getFolderById(attachmentsFolderUrl.match(/folders\/([-\w]+)/)[1]);
          }
          
          if (attachmentsFolder) {
            const messageAttachments = message.getAttachments({ includeInlineImages: false });
            const savedLinks = [];
            
            messageAttachments.forEach(att => {
              // フォルダ内に保存
              const file = attachmentsFolder.createFile(att);
              // アカウント違いでもリンクを知っていれば閲覧できるように共有設定（閲覧権限）を付与
              file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
              savedLinks.push(`📎 ${att.getName()}\n ➡ ${file.getUrl()}`);
            });
            
            if (savedLinks.length > 0) {
              attachmentText = savedLinks.join("\n");
            }
          }
        } catch(attErr) {
          console.error(`⚠️ [添付ファイル保存エラー] ${attErr.message}`);
          attachmentText = `⚠️ 添付ファイルの保存に失敗しました: ${attErr.message}`;
        }
      }

      // メモ列（M列）にセットするテキスト（添付ファイルリンク、無ければ送信元アドレス）
      const memoContent = attachmentText ? `${attachmentText}\n\n送信元: ${from}` : `送信元: ${from}`;

      // ToDoシートの最終行の次に追加
      const lastRow = todoSheet.getLastRow() + 1;
      
      todoSheet.appendRow([
        receptionDate, 'メール', titleFormula, taskContent, taskPic, 
        finalDueDate, analysis.priority || '中', '未着手',
        false, stakeholder, correspondingName, mailLink, memoContent, ''
      ]);

      // アクション列（I列）にチェックボックスを挿入
      todoSheet.getRange(lastRow, Config.TODO_COL.ACTION + 1).insertCheckboxes();

      // 関係者（J列）と氏名（K列）にプルダウンを動的設定
      setTodoRowValidations(todoSheet, lastRow);

      // 書き込みを即座に反映させて次のループの重複判定に含める
      SpreadsheetApp.flush();

      console.log(`🟢 [ToDo登録成功] シートに追加しました: ${taskTitle} (転送先候補: ${stakeholder})`);

      if (finalDueDate && isTask) {
        try {
          addSingleTaskToCalendar(taskTitle, taskPic, finalDueDate, taskContent);
          console.log(`📅 [カレンダー同期成功] 期限: ${finalDueDate}`);
        } catch(calErr) {
          console.error(`⚠️ [カレンダー登録エラー] ${calErr.message}`);
        }
      }

      // 正常に解析・判定が終わった場合のみ、メールを既読にする
      message.markRead();
      console.log(`✔ [既読化完了] ${subject}`);

      // 🏁 大量メール処理時のAPI制限（連続連投）を回避するため、2秒間休憩を入れる
      if (i < threads.length - 1) {
        console.log(`...API制限回避のため2秒間待機します...`);
        Utilities.sleep(2000);
      }
    }
  } catch (e) {
    console.error("【重大エラー発生】: " + e.stack);
    ss.toast("エラーが発生しました: " + e.message, "システム警告⚠️", 10);
  } finally {
    lock.releaseLock();
  }
}

function askGemini(subject, body, modelName) {
  // 🚨 安全装置：本文が長すぎる場合のAPIエラーを防ぐため、最初の2000文字でカット
  const safeBody = body ? body.substring(0, 2000) : "";

  const prompt = `
あなたは学校の教頭を支える超優秀な秘書AIです。
以下のメールから、教頭のアクションが必要なタスクを抽出し、適切な転送先（関係者）を提案してください。

【出力の絶対条件】
1. summary（内容要約）の文中に、可能であれば「○月○日（曜日）までに」という期限を明記してください。
2. dueDate（期限日付）の厳守事項：
   - 期限が明記されている場合のみ「YYYY/MM/DD」形式で出力してください。
   - 期限が不明、または記載がない場合は、必ず「空文字（""）」を出力してください。
   - 「期限の記載なし」「不明」「未定」といった文字列は絶対に入れないでください。
3. pic（担当者）は「教頭」としてください。
4. stakeholder（転送先提案）の判定：
   - メール本文の内容から、どの部署に転送・処理を依頼するのが適切かを予測して指定してください。
   - 選択肢: '進路指導部', '教務部', '生徒指導部', '事務室', '手元で留める（転送不要）'
   - 特に他の部署に転送する必要がない場合、または教頭自身が対応・確認すべき事務連絡の場合は、必ず '手元で留める（転送不要）' としてください。

【メール内容】
件名: ${subject}
本文: ${safeBody}

回答は必ず以下のJSON形式のみで出力してください。余計なマークアップや説明は一切含めないでください。
{
  "isTask": true または false,
  "title": "要約した件名",
  "summary": "教頭は、～までに～する。",
  "dueDate": "YYYY/MM/DD" または "",
  "priority": "高、中、低",
  "pic": "教頭",
  "stakeholder": "転送先提案（上記選択肢から1つ）"
}
`;

  const payload = { "contents": [{ "parts": [{ "text": prompt }] }] };

  try {
    const response = callGeminiWithRotation(payload, modelName);
    const responseCode = response.getResponseCode();
    
    if (responseCode !== 200) {
      console.error(`[Gemini API エラー] ステータスコード: ${responseCode}`);
      return null;
    }
    
    const json = JSON.parse(response.getContentText());
    if (!json.candidates || !json.candidates[0] || !json.candidates[0].content) {
      console.error("[Gemini API エラー] 応答データの構造が不正です。");
      return null;
    }

    let resultText = json.candidates[0].content.parts[0].text;
    resultText = resultText.replace(/```json/g, "").replace(/```/g, "").trim();
    return JSON.parse(resultText);
  } catch (e) {
    console.error(`[askGemini内の例外エラー] ${e.message}`);
    return null;
  }
}

/**
 * 件名から Re: や Fwd: 等の接頭辞を除去して正規化する
 */
function normalizeSubject(sub) {
  if (!sub) return "";
  let norm = sub.toString().trim();
  
  // 繰り返し現れる返信・転送プレフィックスを再帰的にすべて除去する
  let prev;
  do {
    prev = norm;
    norm = norm
      .replace(/^(fwd|re|fw|転送|返信)\s*[:：]\s*/i, "")
      .replace(/^【\s*(転送|返信|fwd|re|fw)\s*】\s*/i, "")
      .replace(/^\[\s*(転送|返信|fwd|re|fw)\s*\]\s*/i, "")
      .trim();
  } while (norm !== prev);
  
  // MLなどのタグ [ML:123] のようなパターンも除去
  norm = norm.replace(/^\[ml:[^\]]+\]\s*/i, "").trim();
  
  return norm;
}

function isDuplicateEmail(sheet, subject, body, dateStr) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  
  const data = sheet.getDataRange().getValues();
  const normSub = normalizeSubject(subject);
  if (!normSub) return false;
  
  // 比較基準日 (今回解析しているメールの判定日時)
  let targetDate;
  try {
    targetDate = new Date(dateStr);
  } catch (e) {
    targetDate = new Date();
  }
  
  // 7日間のミリ秒数 (7 * 24 * 60 * 60 * 1000)
  const SEVEN_DAYS_MS = 7 * 24 * 60 * 60 * 1000;
  
  for (let i = 1; i < data.length; i++) {
    const existingTitle = String(data[i][Config.TODO_COL.TITLE] || "");
    const existingDateVal = data[i][Config.TODO_COL.DATE];
    
    // Hyperlink式からタイトルを抽出
    let cleanTitle = existingTitle;
    const linkMatch = existingTitle.match(/=HYPERLINK\("[^"]*",\s*"([^"]*)"\)/i);
    if (linkMatch) {
      cleanTitle = linkMatch[1];
    }
    cleanTitle = cleanTitle.replace(/^📎\s*/, "");
    
    const normExistingSub = normalizeSubject(cleanTitle);
    
    // 日付のチェック（過去7日以内かどうか）
    let existingDate = null;
    if (existingDateVal instanceof Date) {
      existingDate = existingDateVal;
    } else if (existingDateVal) {
      try {
        existingDate = new Date(existingDateVal);
      } catch(e) {
        existingDate = null;
      }
    }
    
    if (existingDate && !isNaN(existingDate.getTime())) {
      const diffMs = Math.abs(targetDate.getTime() - existingDate.getTime());
      if (diffMs <= SEVEN_DAYS_MS) {
        // 件名の正規化一致で重複と判断
        if (normSub === normExistingSub) {
          return true;
        }
      }
    }
  }
  return false;
}

function ensureMailAddressSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("メルアド");
  if (!sheet) {
    sheet = ss.insertSheet("メルアド");
    sheet.appendRow(["氏名", "メールアドレス", "役職・係等", "備考"]);
    
    // 見出しの装飾
    sheet.getRange("A1:D1").setBackground("#34a853").setFontColor("#ffffff").setFontWeight("bold");
    
    const defaults = [
      ["進路 太郎", "shinro@example.com", "進路指導部", "進路関係のメール転送先"],
      ["教務 次郎", "kyomu@example.com", "教務部", "時間割・教育課程に関する転送先"],
      ["生徒指導 三郎", "seito@example.com", "生徒指導部", "生徒指導・行事に関する転送先"],
      ["事務 職子", "jimu@example.com", "事務室", "予算・施設・一般事務連絡の転送先"]
    ];
    
    defaults.forEach(row => sheet.appendRow(row));
    sheet.autoResizeColumns(1, 4);
    ss.toast("「メルアド」シートを作成し、サンプルデータを登録しました。適切なアドレスに修正してください。", "⚙️ 初期設定");
  }
}

/**
 * ToDo行の関係者（J列）と氏名（K列）に「メルアド」シートベースのプルダウンを設定する
 */
function setTodoRowValidations(sheet, rowIdx) {
  const ss = sheet.getParent();
  const mailAddressSheet = ss.getSheetByName("メルアド");
  
  let rolesList = ["手元で留める（転送不要）"];
  let namesList = [];
  
  if (mailAddressSheet) {
    const data = mailAddressSheet.getDataRange().getValues();
    const rolesSet = new Set();
    const namesSet = new Set();
    
    for (let i = 1; i < data.length; i++) {
      const name = String(data[i][0] || '').trim(); // A列
      const role = String(data[i][2] || '').trim(); // C列
      if (name) namesSet.add(name);
      if (role) rolesSet.add(role);
    }
    
    if (rolesSet.size > 0) {
      rolesList = Array.from(rolesSet).concat(["手元で留める（転送不要）"]);
    }
    if (namesSet.size > 0) {
      namesList = Array.from(namesSet);
    }
  } else {
    rolesList = ["進路指導部", "教務部", "生徒指導部", "事務室", "手元で留める（転送不要）"];
    namesList = ["進路 太郎", "教務 次郎", "生徒指導 三郎", "事務 職子"];
  }
  
  // J列 (関係者/役職) の入力規則
  const roleRule = SpreadsheetApp.newDataValidation().requireValueInList(rolesList, true).setAllowInvalid(true).build();
  sheet.getRange(rowIdx, Config.TODO_COL.STAKEHOLDER + 1).setDataValidation(roleRule);
  
  // K列 (氏名) の入力規則
  if (namesList.length > 0) {
    const nameRule = SpreadsheetApp.newDataValidation().requireValueInList(namesList, true).setAllowInvalid(true).build();
    sheet.getRange(rowIdx, Config.TODO_COL.NAME + 1).setDataValidation(nameRule);
  }
}

/**
 * 役職名から「メルアド」シートを検索し、対応する氏名を返す
 */
function getNameByRole(role) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("メルアド");
  if (!sheet || !role) return null;
  
  const data = sheet.getDataRange().getValues();
  const roles = role.split(/[,，、\n\r]+/).filter(r => r.trim() !== "");
  const names = [];
  
  roles.forEach(r => {
    for (let i = 1; i < data.length; i++) {
      const dbRole = String(data[i][2] || '').trim(); // C列: 役職
      const dbName = String(data[i][0] || '').trim(); // A列: 氏名
      if (dbRole === r && dbName) {
        names.push(dbName);
        break;
      }
    }
  });
  
  return names.length > 0 ? names.join(", ") : null;
}

/**
 * 氏名から「メルアド」シートを検索し、対応する役職名を返す
 */
function getRoleByName(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("メルアド");
  if (!sheet || !name) return null;
  
  const data = sheet.getDataRange().getValues();
  const names = name.split(/[,，、\n\r]+/).filter(n => n.trim() !== "");
  const roles = [];
  
  names.forEach(n => {
    for (let i = 1; i < data.length; i++) {
      const dbName = String(data[i][0] || '').trim(); // A列: 氏名
      const dbRole = String(data[i][2] || '').trim(); // C列: 役職
      if (dbName === n && dbRole) {
        roles.push(dbRole);
        break;
      }
    }
  });
  
  return roles.length > 0 ? roles.join(", ") : null;
}

/**
 * 氏名または役職から転送先メールアドレスを検索・取得する
 */
function getForwardEmail(stakeholder, name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("メルアド");
  if (!sheet) return null;
  
  const data = sheet.getDataRange().getValues();
  const emails = [];
  
  // 1. 氏名（K列）の指定があれば、複数氏名を分割してメールアドレスを検索
  if (name) {
    const names = name.split(/[,，、\n\r]+/).filter(n => n.trim() !== "");
    names.forEach(n => {
      for (let i = 1; i < data.length; i++) {
        const dbName = String(data[i][0] || '').trim();
        const email = String(data[i][1] || '').trim();
        if (dbName === n && email) {
          emails.push(email);
          break;
        }
      }
    });
  }
  
  // 2. 氏名で見つからなかったか、氏名指定がない場合に、役職（C列）ベースでメールアドレスを検索
  if (emails.length === 0 && stakeholder) {
    const roles = stakeholder.split(/[,，、\n\r]+/).filter(r => r.trim() !== "");
    roles.forEach(r => {
      let found = false;
      for (let i = 1; i < data.length; i++) {
        const dbRole = String(data[i][2] || '').trim();
        const email = String(data[i][1] || '').trim();
        if (dbRole === r && email) {
          emails.push(email);
          found = true;
          break;
        }
      }
      if (!found) {
        const configDest = Config.FORWARD_DESTINATIONS;
        if (configDest && configDest[r]) {
          emails.push(configDest[r]);
        }
      }
    });
  }
  
  // 重複を除去してカンマ区切りで結合
  const uniqueEmails = Array.from(new Set(emails));
  return uniqueEmails.length > 0 ? uniqueEmails.join(", ") : null;
}

/**
 * アクションチェックボックスがONにされた際のメール転送実行およびタスク消し込み
 */
function handleTodoAction(sheet, rowIdx) {
  const ss = sheet.getParent();
  try {
    const rowValues = sheet.getRange(rowIdx, 1, 1, sheet.getLastColumn()).getValues()[0];
    const status = rowValues[Config.TODO_COL.STATUS];
    
    if (status === '完了') {
      ss.toast('このタスクはすでに完了しています。', '⚠️ 警告');
      return;
    }
    
    const titleFormula = rowValues[Config.TODO_COL.TITLE];
    const source = rowValues[Config.TODO_COL.SOURCE];
    const stakeholder = String(rowValues[Config.TODO_COL.STAKEHOLDER] || '').trim();
    const name = String(rowValues[Config.TODO_COL.NAME] || '').trim();
    
    let mailLink = rowValues[Config.TODO_COL.MAIL_LINK];
    const memoVal = rowValues[Config.TODO_COL.MEMO];
    // 過去の行（列ズレ）対策：L列が空で、M列（メモ）にGmailリンクがある場合はそれを使用する
    if ((!mailLink || mailLink.toString().trim() === "") && memoVal && memoVal.toString().includes('mail.google.com')) {
      mailLink = memoVal;
    }
    
    // ハイパーリンク式や添付アイコンを除去した純粋なタイトルを取得
    let title = titleFormula;
    const linkMatch = titleFormula.match(/=HYPERLINK\("[^"]*",\s*"([^"]*)"\)/i);
    if (linkMatch) {
      title = linkMatch[1];
    }
    title = title.replace(/^📎\s*/, "");
    
    if (source !== 'メール') {
      // メール以外の場合は、単純に完了ステータスに変更
      sheet.getRange(rowIdx, Config.TODO_COL.STATUS + 1).setValue('完了');
      sheet.getRange(rowIdx, Config.TODO_COL.COMPLETED_DATE + 1).setValue(new Date());
      ss.toast(`タスク「${title}」を完了にしました。`, '✅ 完了');
      return;
    }
    
    if (stakeholder === '手元で留める（転送不要）' || (!stakeholder && !name)) {
      sheet.getRange(rowIdx, Config.TODO_COL.STATUS + 1).setValue('完了');
      sheet.getRange(rowIdx, Config.TODO_COL.COMPLETED_DATE + 1).setValue(new Date());
      ss.toast(`手元で留める（転送不要）として処理を完了しました。`, '✅ 完了');
      return;
    }
    
    // 転送先メールアドレスを検索（氏名優先 ➡ 役職）
    const recipientEmail = getForwardEmail(stakeholder, name);
    if (!recipientEmail) {
      const targetLabel = name ? `${name}先生` : `「${stakeholder}」`;
      SpreadsheetApp.getUi().alert(`エラー: ${targetLabel} の転送先メールアドレスが「メルアド」シートに見つかりません。`);
      return;
    }
    
    // メールリンクからメッセージIDを抽出
    let messageId = "";
    if (mailLink) {
      const mailLinkStr = mailLink.toString().trim();
      
      // 1. 正規表現で16文字以上の16進数（GmailのメッセージID）をパースする
      const match = mailLinkStr.match(/\/([0-9a-fA-F]{16,})\b/);
      if (match) {
        messageId = match[1];
      } else {
        // 2. フォールバック: #inbox/ や #all/ などの後ろをパース
        const hashMatch = mailLinkStr.match(/#(?:inbox|all|sent|drafts|trash|archive|label\/[^/]+)\/([0-9a-fA-F]+)/i);
        if (hashMatch) {
          messageId = hashMatch[1];
        } else {
          // 3. 最終手段: スラッシュで分割して最後のセグメントを取得
          const parts = mailLinkStr.split('/');
          const lastSegment = parts[parts.length - 1].split('?')[0].trim();
          if (/^[0-9a-fA-F]+$/.test(lastSegment)) {
            messageId = lastSegment;
          }
        }
      }
    }
    
    if (!messageId) {
      SpreadsheetApp.getUi().alert(
        '⚠️ 転送エラー',
        `GmailのメッセージIDを取得できませんでした。\n\n` +
        `【原因の可能性】\n` +
        `1. メールから自動登録されたToDoではない可能性があります。（手動登録、音声メモ、電話伝言などには転送元メールがありません）\n` +
        `2. スプレッドシートの列がズレているか、セルの値が書き換わっている可能性があります。\n\n` +
        `【セルの値】\nメールリンク列（L列）の値: "${mailLink || '（空欄）'}"`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    const targetLabel = name ? `${name}先生` : stakeholder;
    
    // ✨ 誤送信防止のための確認ダイアログを表示
    const ui = SpreadsheetApp.getUi();
    const confirmResponse = ui.alert(
      '📧 メール転送の確認',
      `以下の宛先にメールを転送しますか？\n\n宛先: ${targetLabel} (${recipientEmail})\n件名: ${title}`,
      ui.ButtonSet.YES_NO
    );
    if (confirmResponse !== ui.Button.YES) {
      ss.toast('転送をキャンセルしました。', 'ℹ️ キャンセル');
      return;
    }
    
    ss.toast(`メールを「${targetLabel}」(${recipientEmail})へ転送中...`, '📤 転送処理中');
    
    // Gmailメッセージを取得して転送
    let message;
    try {
      message = GmailApp.getMessageById(messageId);
    } catch (err) {
      SpreadsheetApp.getUi().alert(
        '⚠️ メールが見つかりません',
        `Gmailのメッセージ（ID: ${messageId}）を取得できませんでした。\n\n` +
        `【原因の可能性】\n` +
        `1. 現在ログインしているGoogleアカウントのメールボックスに、このメールが存在しない可能性があります。（別のアカウントで作成されたToDo行をコピーした場合など）\n` +
        `2. 対象のメールがGmail側ですでに完全に削除（ゴミ箱からも消去）されている可能性があります。\n\n` +
        `【実際のメールリンクURL】\n"${mailLink || '（空欄）'}"`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    const attachments = message.getAttachments({ includeInlineImages: true });
    
    message.forward(recipientEmail, {
      subject: `Fwd: ${message.getSubject()}`,
      attachments: attachments
    });
    
    // 状態と完了日を更新
    sheet.getRange(rowIdx, Config.TODO_COL.STATUS + 1).setValue('完了');
    sheet.getRange(rowIdx, Config.TODO_COL.COMPLETED_DATE + 1).setValue(new Date());
    
    ss.toast(`「${targetLabel}」へメールを転送し、タスクを完了にしました！`, '✅ 転送完了');
    
  } catch(e) {
    console.error("handleTodoAction error:", e);
    SpreadsheetApp.getUi().alert(`転送アクションの実行中にエラーが発生しました:\n${e.toString()}`);
  }
}

/**
 * ToDoシートの列ヘッダー構造を強制的に定義通りに修正し、
 * K列（氏名）とL列（メールリンク）のデータのテレコ（逆転ズレ）や、
 * I列（アクション）のチェックボックス欠落を自動で修復する
 */
function ensureTodoSheetColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(Config.SHEET_NAME_TODO);
  if (!sheet) return;
  
  // 1. 各列のヘッダー名（1〜14列目）を定義通りに強制リセットしてズレを防ぐ
  const expectedHeaders = [
    "受信日時", "発生源", "件名", "内容（詳細）", "担当", "期限", "重要度", "ステータス",
    "処理を実行", "転送先（部署）", "氏名", "メールリンク", "メモ", "完了日"
  ];
  
  // シートの現在の列数を取得し、15列未満なら一時的に15列まで拡張する（getRangeのエラーを防ぐため）
  let currentLastCol = sheet.getLastColumn();
  if (currentLastCol < 15) {
    sheet.insertColumnsAfter(currentLastCol, 15 - currentLastCol);
    currentLastCol = 15;
  }
  
  // ヘッダーを強制上書き
  sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // 2. ✨ データ列のズレ自動修復（2次元配列上での一括バッチ処理）
  // 2行目から最終行までの15列分を一括で取得
  const range = sheet.getRange(2, 1, lastRow - 1, 15);
  const data = range.getValues();
  let fixCount = 0;
  
  for (let i = 0; i < data.length; i++) {
    // インデックス定義
    // K列 (氏名): インデックス10 (11列目)
    // L列 (メールリンク): インデックス11 (12列目)
    // M列 (メモ): インデックス12 (13列目)
    // N列 (完了日): インデックス13 (14列目)
    // O列 (予備・2重完了日): インデックス14 (15列目)
    
    const kVal = String(data[i][10] || '').trim();
    const lVal = String(data[i][11] || '').trim();
    const mVal = String(data[i][12] || '').trim();
    const nVal = data[i][13];
    const oVal = data[i][14];
    
    // 判定パターンA: M列（インデックス12）にURL（http）が含まれている場合（移行期の主要な列ズレ）
    if (mVal.includes('mail.google.com')) {
      // a. もしL列に名前が入っている場合は、K列に氏名を移動
      if (!kVal && lVal && !lVal.includes('http')) {
        data[i][10] = lVal;
      }
      
      // b. M列のURLをL列へ
      data[i][11] = mVal;
      
      // c. N列のメモをM列へ
      data[i][12] = nVal;
      
      // d. O列の完了日をN列へ
      data[i][13] = oVal;
      
      // e. O列（インデックス14）はクリア
      data[i][14] = "";
      
      fixCount++;
    }
    // 判定パターンB: K列（インデックス10）にURLがあり、L列（インデックス11）に名前があるテレコ（スワップ）状態
    else if (kVal.includes('http') && !lVal.includes('http')) {
      data[i][10] = lVal; // K列に名前
      data[i][11] = kVal; // L列にURL
      fixCount++;
    }
  }
  
  // 修正後のデータを一括でシートに書き戻す
  if (fixCount > 0) {
    range.setValues(data);
    ss.toast("過去の古いToDo行の列ズレ（" + fixCount + "件）を自動修復しました。", "⚙️ 自動メンテナンス", 4000);
    SpreadsheetApp.flush();
  }

  // 3. 余分な15列目（O列）以降を完全に削除して綺麗にする
  const finalLastCol = sheet.getLastColumn();
  if (finalLastCol >= 15) {
    // 15列目から始まる余分な列を削除
    sheet.deleteColumns(15, finalLastCol - 15 + 1);
    ss.toast("重複していた不要な「完了日」（O列）を削除しました。", "⚙️ 整理完了", 3000);
    SpreadsheetApp.flush();
  }

  // 4. 各種プルダウン・データの入力規則を一括設定し直す
  // L列（メールリンク）、M列（メモ）、N列（完了日）の不要な古い入力規則（プルダウン等）をクリア
  sheet.getRange(2, Config.TODO_COL.MAIL_LINK + 1, lastRow - 1, 3).clearDataValidations();
  
  // メルアドリストの一括取得と入力規則の作成
  const mailAddressSheet = ss.getSheetByName("メルアド");
  let rolesList = ["手元で留める（転送不要）"];
  let namesList = [];
  
  if (mailAddressSheet) {
    const maData = mailAddressSheet.getDataRange().getValues();
    const rolesSet = new Set();
    const namesSet = new Set();
    for (let i = 1; i < maData.length; i++) {
      const name = String(maData[i][0] || '').trim();
      const role = String(maData[i][2] || '').trim();
      if (name) namesSet.add(name);
      if (role) rolesSet.add(role);
    }
    if (rolesSet.size > 0) rolesList = Array.from(rolesSet).concat(["手元で留める（転送不要）"]);
    if (namesSet.size > 0) namesList = Array.from(namesSet);
  } else {
    rolesList = ["進路指導部", "教務部", "生徒指導部", "事務室", "手元で留める（転送不要）"];
    namesList = ["進路 太郎", "教務 次郎", "生徒指導 三郎", "事務 職子"];
  }
  
  const roleRule = SpreadsheetApp.newDataValidation().requireValueInList(rolesList, true).setAllowInvalid(true).build();
  sheet.getRange(2, Config.TODO_COL.STAKEHOLDER + 1, lastRow - 1, 1).setDataValidation(roleRule);
  
  if (namesList.length > 0) {
    const nameRule = SpreadsheetApp.newDataValidation().requireValueInList(namesList, true).setAllowInvalid(true).build();
    sheet.getRange(2, Config.TODO_COL.NAME + 1, lastRow - 1, 1).setDataValidation(nameRule);
  }

  // 5. --- ✨ 古いToDo行のI列（アクション）にチェックボックスが無い場合の自動復旧処理 ---
  const actionColRange = sheet.getRange(2, Config.TODO_COL.ACTION + 1, lastRow - 1, 1);
  const actionValues = actionColRange.getValues();
  const validations = actionColRange.getDataValidations();
  let checkboxRecoverCount = 0;
  
  for (let i = 0; i < actionValues.length; i++) {
    const val = actionValues[i][0];
    const validation = validations[i][0];
    
    // データの検証が設定されていない、またはチェックボックスでない場合
    if (!validation || validation.getCriteriaType() !== SpreadsheetApp.DataValidationCriteria.CHECKBOX) {
      const rowIdx = i + 2;
      
      // 文字列が入っている場合は、退避してメモ欄に書き出す
      if (typeof val === 'string' && val.trim() !== '' && val !== 'false' && val !== 'true') {
        const memoRange = sheet.getRange(rowIdx, Config.TODO_COL.MEMO + 1);
        const currentMemo = memoRange.getValue().toString().trim();
        const newMemo = currentMemo ? currentMemo + " (元アクション: " + val + ")" : "【元アクション】" + val;
        memoRange.setValue(newMemo);
      }
      
      // チェックボックスを挿入してデフォルトをfalseにする
      sheet.getRange(rowIdx, Config.TODO_COL.ACTION + 1).insertCheckboxes().setValue(false);
      checkboxRecoverCount++;
    }
  }
  
  if (checkboxRecoverCount > 0) {
    ss.toast("古いToDoのチェックボックスを自動復旧しました。", "⚙️ 自動メンテナンス");
  }
}
