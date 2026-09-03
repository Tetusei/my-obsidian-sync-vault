/**
 * 次回にそなえた片づけ（リセット）。
 *
 * 消す範囲を3段階に分けている。
 *   1. 予約・だめなコマ・ログを個別に消す … やり直しや練習後の片づけ
 *   2. 次の面談にそなえて初期化      … 名簿と担任はそのまま、予約だけ白紙に
 *   3. 年度末：すべて初期化           … 名簿も担任も面談日も消す
 *
 * 2と3は取り返しがつかないので、実行前に必ずバックアップを取り、
 * 決められた言葉を入力してもらってから進める。
 */

/* ---------------- メニュー処理 ---------------- */

function menuClearBookings() {
  var ui = SpreadsheetApp.getUi();
  try {
    var n = countBookings_();
    if (!n) { ui.alert('取り消す予約はありません。'); return; }

    var res = ui.alert('予約をすべて取り消しますか？',
      '現在 ' + n + ' 件の予約が入っています。すべて取り消して空き枠に戻します。\n' +
      '（「' + SH.NG + '」で指定した枠は、空きではなく「面談なし」に戻ります）\n\n' +
      '面談日・名簿・だめなコマの指定はそのまま残ります。\n' +
      '保護者への連絡は行われません。取り消したことは別途お伝えください。',
      ui.ButtonSet.OK_CANCEL);
    if (res !== ui.Button.OK) return;

    var cleared = clearAllBookings_();
    refreshAfterReset_();
    ui.alert('完了', cleared + ' 件の予約を取り消しました。', ui.ButtonSet.OK);
  } catch (err) {
    ui.alert('エラー', String(err.message || err), ui.ButtonSet.OK);
  }
}

function menuClearNg() {
  var ui = SpreadsheetApp.getUi();
  try {
    var n = countNgMarks_();
    if (!n) { ui.alert('だめなコマの指定はありません。'); return; }

    var res = ui.alert('だめなコマの指定をすべて外しますか？',
      '現在 ' + n + ' 枠が「面談を入れない」に指定されています。\n' +
      'すべてチェックを外し、空き枠に戻します。予約は消えません。',
      ui.ButtonSet.OK_CANCEL);
    if (res !== ui.Button.OK) return;

    var cleared = clearAllNg_();
    var applied = applyNgSlots();
    refreshAfterReset_();
    ui.alert('完了',
      cleared + ' 枠の指定を外し、' + applied.unblocked + ' 枠を空きに戻しました。',
      ui.ButtonSet.OK);
  } catch (err) {
    ui.alert('エラー', String(err.message || err), ui.ButtonSet.OK);
  }
}

function menuClearLog() {
  var ui = SpreadsheetApp.getUi();
  try {
    var n = countLog_();
    if (!n) { ui.alert('予約ログは空です。'); return; }

    var res = ui.alert('予約ログを消しますか？',
      n + ' 行の履歴を削除します。予約そのものには影響しません。\n' +
      '問い合わせ対応の記録がなくなるため、面談が終わってから実行してください。',
      ui.ButtonSet.OK_CANCEL);
    if (res !== ui.Button.OK) return;

    var cleared = clearLog_();
    ui.alert('完了', cleared + ' 行の履歴を削除しました。', ui.ButtonSet.OK);
  } catch (err) {
    ui.alert('エラー', String(err.message || err), ui.ButtonSet.OK);
  }
}

/** 次の面談にそなえて初期化（名簿・担任・設定の時間類は残す） */
function menuResetForNext() {
  var ui = SpreadsheetApp.getUi();
  try {
    var summary =
      '【対象のファイル】' + LF + ss_().getName() + LF + LF +
      '次の面談にそなえて、予約まわりを白紙に戻します。\n\n' +
      '【消えるもの】\n' +
      '・予約 ' + countBookings_() + ' 件（すべて取り消し）\n' +
      '・だめなコマの指定 ' + countNgMarks_() + ' 枠\n' +
      '・予約ログ ' + countLog_() + ' 行\n' +
      '・予約受付の開始日時・締切日時\n\n' +
      '【残るもの】\n' +
      '・生徒名簿 ' + countRoster_() + ' 名（' + CLASS_SHEET_PREFIX + '〇組 のA・B列）\n' +
      '・クラス名と担任\n' +
      '・面談の時間設定（開始時刻・枠の長さ・枠数）\n' +
      '・面談日（次回の日付に書き換えてお使いください）\n\n' +
      '実行前に自動でバックアップを取ります。';

    if (!confirmTyped_(ui, '次の面談にそなえて初期化', summary, 'リセット')) return;

    ss_().toast('バックアップを取っています…', '三者面談 リセット', 10);
    var backup = createManualBackup();

    var result = resetForNext_();
    refreshAfterReset_();

    ui.alert('初期化しました',
      '予約 ' + result.bookings + ' 件、だめなコマ ' + result.ng + ' 枠、ログ ' + result.log + ' 行を消しました。\n' +
      '予約受付は停止状態になっています。\n\n' +
      '【次の手順】\n' +
      '1. 「' + SH.DAYS + '」シートの日付を次回のものに書き換える\n' +
      '2. 「' + MENU.GENERATE + '」を実行\n' +
      '3. 「' + MENU.NG_SHEET + '」→ 面談を入れない枠にチェック →「' + MENU.NG_APPLY + '」\n' +
      '4. 「' + MENU.PUBLISH + '」\n\n' +
      'バックアップ: ' + backup.fileName,
      ui.ButtonSet.OK);
  } catch (err) {
    ui.alert('エラー', String(err.message || err), ui.ButtonSet.OK);
  }
}

/** 年度末：すべて初期化（名簿・担任・面談日も消す） */
function menuResetAll() {
  var ui = SpreadsheetApp.getUi();
  try {
    var summary =
      '【対象のファイル】' + LF + ss_().getName() + LF + LF +
      '年度末の片づけです。名簿を含めて、この学年のデータをすべて消します。\n\n' +
      '【消えるもの】\n' +
      '・生徒名簿 ' + countRoster_() + ' 名\n' +
      '・担任名と担任メール（クラス名は残ります）\n' +
      '・面談日\n' +
      '・予約 ' + countBookings_() + ' 件、面談枠すべて\n' +
      '・だめなコマの指定 ' + countNgMarks_() + ' 枠\n' +
      '・予約ログ ' + countLog_() + ' 行\n\n' +
      '【残るもの】\n' +
      '・クラス名\n' +
      '・面談の時間設定と管理パスコード\n' +
      '・作成済みのPDFとバックアップ（別フォルダのため消えません）\n\n' +
      '実行前に自動でバックアップを取ります。元に戻したいときはそこから復元してください。';

    if (!confirmTyped_(ui, '⚠ 年度末：すべて初期化', summary, '全初期化')) return;

    ss_().toast('バックアップを取っています…', '三者面談 リセット', 10);
    var backup = createManualBackup();

    var result = resetAll_();
    refreshAfterReset_();

    ui.alert('すべて初期化しました',
      '名簿 ' + result.roster + ' 名、予約 ' + result.bookings + ' 件、面談枠 ' + result.slots + ' 枠を消しました。\n\n' +
      '【次年度の手順】\n' +
      '1. 「' + SH.CLASSES + '」シートに担任名とメールを入力\n' +
      '2. 各「' + CLASS_SHEET_PREFIX + '〇組」シートのA・B列に新しい名簿を貼り付け\n' +
      '3. 「' + SH.DAYS + '」シートに面談日を入力\n' +
      '4. 「' + MENU.GENERATE + '」を実行\n\n' +
      'バックアップ: ' + backup.fileName,
      ui.ButtonSet.OK);
  } catch (err) {
    ui.alert('エラー', String(err.message || err), ui.ButtonSet.OK);
  }
}

/**
 * 他校へ配るためのひな形にする。
 *
 * 年度末の全初期化では、URL・管理パスコード・自動処理の設定が残る。
 * 「保護者用URL」が残ったまま配ると、コピーした学校の案内プリントのQRコードが
 * こちらの予約ページを指してしまい、他校の保護者がこちらに予約してしまう。
 * 配布の前に、この学校に固有のものをすべて落とす。
 */
function menuResetForDistribution() {
  var ui = SpreadsheetApp.getUi();
  try {
    var summary =
      '【対象のファイル】' + LF + ss_().getName() + LF + LF +
      'このファイルを、他校へ配る「ひな形」にします。' + LF +
      '年度末の全初期化に加えて、この学校に固有の情報をすべて消します。' + LF + LF +
      '【消えるもの】' + LF +
      '・生徒名簿 ' + countRoster_() + ' 名、担任名とメール' + LF +
      '・面談日、予約 ' + countBookings_() + ' 件、面談枠すべて' + LF +
      '・だめなコマの指定、予約ログ' + LF +
      '・保護者用URL と 管理画面URL' + LF +
      '・管理パスコード' + LF +
      '・自動処理（表示の更新・バックアップ・リマインド・だめなコマ）のトリガー' + LF + LF +
      '【残るもの】' + LF +
      '・クラス名' + LF +
      '・面談の時間設定（開始時刻・枠の長さ・休憩・枠数）' + LF + LF +
      '⚠ 保護者用URLを消さずに配ると、コピーした学校の案内プリントのQRコードが' + LF +
      '　 この学校の予約ページを指してしまいます。必ずこの操作で消してください。' + LF + LF +
      '実行前に自動でバックアップを取ります。';

    if (!confirmTyped_(ui, '⚠ 配布用にまっさらにする', summary, '配布用')) return;

    ss_().toast('バックアップを取っています…', '三者面談 リセット', 10);
    var backup = createManualBackup();

    var result = resetAll_();
    var stopped = clearSchoolIdentity_();
    refreshAfterReset_();

    ui.alert('配布用に初期化しました',
      '名簿 ' + result.roster + ' 名、予約 ' + result.bookings + ' 件、面談枠 ' + result.slots + ' 枠を消しました。' + LF +
      'URL・管理パスコードを空にし、自動処理 ' + stopped + ' 件を止めました。' + LF + LF +
      '【配る前に、目で確かめてください】' + LF +
      '1. 「' + SH.CONFIG + '」の 保護者用URL・管理画面URL・管理パスコード が空欄' + LF +
      '2. 各「' + CLASS_SHEET_PREFIX + '〇組」のA・B列が空' + LF +
      '3. 「' + SH.DAYS + '」が空' + LF +
      '4. 「公開」が FALSE' + LF +
      '5. Apps Script の「デプロイを管理」で、公開中のデプロイをアーカイブ' + LF + LF +
      'バックアップ: ' + backup.fileName,
      ui.ButtonSet.OK);
  } catch (err) {
    ui.alert('エラー', String(err.message || err), ui.ButtonSet.OK);
  }
}

/**
 * この学校に固有のもの（URL・合言葉・自動処理）を消す。
 * トリガーはコピー先に引き継がれないが、ひな形側で動き続ける意味がない。
 * @return {number} 止めたトリガーの数
 */
function clearSchoolIdentity_() {
  setConfigValue_('保護者用URL', '');
  setConfigValue_(ADMIN_URL_KEY, '');
  setConfigValue_('管理パスコード', '');

  var stopped = 0;
  var offs = [disableAutoRefresh, disableNgAutoApply, disableReminder, disableAutoBackup];
  for (var i = 0; i < offs.length; i++) {
    try {
      stopped += Number(offs[i]()) || 0;
    } catch (e) {
      console.warn('自動処理を止められませんでした:', e);
    }
  }
  return stopped;
}

/* ---------------- 実処理 ---------------- */

/**
 * 次の面談にそなえた初期化。
 * @return {{bookings:number, ng:number, log:number}}
 */
function resetForNext_() {
  var bookings = clearAllBookings_();
  var ng = clearAllNg_();
  applyNgSlots();                       // ブロックを空きに戻す
  var log = clearLog_();

  setConfigValue_('公開', false);
  setConfigValue_('予約受付開始', '');
  setConfigValue_('予約受付締切', '');

  logAction_('リセット', '', '', '', '',
    '次回にそなえて初期化: 予約 ' + bookings + '件 / だめなコマ ' + ng + '枠 / ログ ' + log + '行');

  return { bookings: bookings, ng: ng, log: log };
}

/**
 * 年度末の全初期化。
 * @return {{bookings:number, ng:number, log:number, roster:number, slots:number}}
 */
function resetAll_() {
  var base = resetForNext_();
  var roster = clearRoster_();
  clearTeachers_();
  clearDays_();
  var slots = clearSlots_();
  clearNgSheet_();
  clearLog_();                          // resetForNext_ 以降の追記も消す

  return {
    bookings: base.bookings,
    ng: base.ng,
    log: base.log,
    roster: roster,
    slots: slots
  };
}

/**
 * 予約をすべて取り消す。だめなコマに指定された枠は「面談なし」に戻す。
 * @return {number} 取り消した件数
 */
function clearAllBookings_() {
  return withLock_(function () {
    var sh = sheet_(SH.SLOTS);
    var last = sh.getLastRow();
    if (last < 2) return 0;

    var ng = {};
    try { ng = readNgSet_(); } catch (e) { ng = {}; }

    var width = SLOT_LAST_COL - COL.STATUS + 1;
    var slots = readSlots_();
    var block = sh.getRange(2, COL.STATUS, last - 1, width).getValues();
    var count = 0;

    for (var i = 0; i < slots.length; i++) {
      var v = slots[i].v;
      var st = String(v[COL.STATUS - 1]);

      // 予備の枠は枠そのものを残し、入れた生徒だけを消す
      if (st === STATUS.RESERVE) {
        if (!v[COL.NUMBER - 1] && !v[COL.STUDENT - 1]) continue;
        block[slots[i].row - 2] = [STATUS.RESERVE, '', '', '', '', '', ''];
        count++;
        continue;
      }

      if (st !== STATUS.BOOKED) continue;
      var id = String(v[COL.SLOT_ID - 1]);
      block[slots[i].row - 2] = [ng[id] ? STATUS.BLOCKED : STATUS.OPEN, '', '', '', '', '', ''];
      count++;
    }

    if (count) {
      sh.getRange(2, COL.STATUS, last - 1, width).setValues(block);
      clearSlotCache_();
      logAction_('一括取消', '', '', '', '', count + '件の予約を取り消しました');
    }
    return count;
  });
}

/**
 * だめなコマのチェックをすべて外す。
 * @return {number} 外した数
 */
function clearAllNg_() {
  var sh = ss_().getSheetByName(SH.NG);
  if (!sh) return 0;

  var last = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (last < NG_FIRST_ROW || lastCol < NG_COL.FIRST_CLASS) return 0;

  var rowCount = last - NG_FIRST_ROW + 1;
  var width = lastCol - NG_COL.FIRST_CLASS + 1;
  var rng = sh.getRange(NG_FIRST_ROW, NG_COL.FIRST_CLASS, rowCount, width);
  var vals = rng.getValues();

  var count = 0;
  var out = [];
  for (var r = 0; r < rowCount; r++) {
    var row = [];
    for (var c = 0; c < width; c++) {
      if (isNgMark_(vals[r][c])) count++;
      row.push(false);
    }
    out.push(row);
  }
  if (count) rng.setValues(out);
  return count;
}

/** 予約ログの中身を消す（見出しは残す） */
function clearLog_() {
  var sh = ss_().getSheetByName(SH.LOG);
  if (!sh) return 0;
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var cols = Math.max(sh.getLastColumn(), 1);
  sh.getRange(2, 1, last - 1, cols).clearContent();
  return last - 1;
}

/** 各「予約表_〇組」シートのA・B列（名簿）を消す */
function clearRoster_() {
  var ss = ss_();
  var classes = getClasses();
  var count = 0;

  for (var c = 0; c < classes.length; c++) {
    var sh = ss.getSheetByName(CLASS_SHEET_PREFIX + classes[c].name);
    if (!sh) continue;
    var last = sh.getLastRow();
    if (last < 2) continue;

    var vals = sh.getRange(2, 1, last - 1, 2).getValues();
    for (var i = 0; i < vals.length; i++) {
      if (vals[i][0] && vals[i][1]) count++;
    }
    sh.getRange(2, 1, last - 1, 2).clearContent();
  }
  clearRosterCache_();
  return count;
}

/** 「クラス」シートの担任名・担任メールを消す（クラス名は残す） */
function clearTeachers_() {
  var sh = ss_().getSheetByName(SH.CLASSES);
  if (!sh) return 0;
  var last = sh.getLastRow();
  if (last < 2) return 0;
  sh.getRange(2, 2, last - 1, 2).clearContent();
  dropRefCaches_();
  return last - 1;
}

/** 「面談日」シートの中身を消す */
function clearDays_() {
  var sh = ss_().getSheetByName(SH.DAYS);
  if (!sh) return 0;
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var cols = Math.max(sh.getLastColumn(), 3);
  sh.getRange(2, 1, last - 1, cols).clearContent();
  return last - 1;
}

/** 「枠マスタ」シートの中身を消す */
function clearSlots_() {
  var sh = ss_().getSheetByName(SH.SLOTS);
  if (!sh) return 0;
  var last = sh.getLastRow();
  if (last < 2) return 0;
  sh.getRange(2, 1, last - 1, SLOT_LAST_COL).clearContent();
  clearSlotCache_();
  return last - 1;
}

/** 「だめなコマ」シートの表を空にする（見出しとバナーは残す） */
function clearNgSheet_() {
  var sh = ss_().getSheetByName(SH.NG);
  if (!sh) return 0;
  var last = sh.getLastRow();
  var lastCol = Math.max(sh.getLastColumn(), NG_COL.FIRST_CLASS);
  if (last < NG_FIRST_ROW) return 0;

  var rowCount = last - NG_FIRST_ROW + 1;
  var rng = sh.getRange(NG_FIRST_ROW, 1, rowCount, lastCol);
  rng.clearDataValidations();
  rng.clearContent();
  rng.setNotes(emptyGrid_(rowCount, lastCol, ''));
  rng.setBackgrounds(emptyGrid_(rowCount, lastCol, NG_NORMAL_BG));
  setNgBanner_(sh, lastCol, []);
  return rowCount;
}

function emptyGrid_(rows, cols, value) {
  var out = [];
  for (var r = 0; r < rows; r++) {
    var row = [];
    for (var c = 0; c < cols; c++) row.push(value);
    out.push(row);
  }
  return out;
}

/* ---------------- 補助 ---------------- */

/** 決められた言葉を入力してもらう確認。入力が違えば false。 */
function confirmTyped_(ui, title, message, word) {
  var res = ui.prompt(title,
    message + '\n\n続けるには、下の欄に「' + word + '」と入力して［OK］を押してください。',
    ui.ButtonSet.OK_CANCEL);

  if (res.getSelectedButton() !== ui.Button.OK) return false;
  if (norm_(res.getResponseText()) !== norm_(word)) {
    ui.alert('中止しました', '入力された言葉が違うため、何も変更していません。', ui.ButtonSet.OK);
    return false;
  }
  return true;
}

/** リセット後の表示更新。シートが揃っていない場合もあるので個別に握りつぶす。 */
function refreshAfterReset_() {
  try { rebuildOverview(); } catch (e) { console.warn('全体ビューの更新をスキップ:', e); }
  try { rebuildClassSheets(); } catch (e) { console.warn('クラス別予約表の更新をスキップ:', e); }
  try { hideInternalSheets(); } catch (e) { /* 無視 */ }
}

function countBookings_() {
  var n = 0;
  try {
    var slots = readSlots_();
    for (var i = 0; i < slots.length; i++) {
      if (isTakenSlot_(slots[i].v)) n++;
    }
  } catch (e) { /* 枠マスタが無ければ0 */ }
  return n;
}

function countNgMarks_() {
  try { return Object.keys(readNgSet_()).length; } catch (e) { return 0; }
}

function countLog_() {
  var sh = ss_().getSheetByName(SH.LOG);
  return sh && sh.getLastRow() > 1 ? sh.getLastRow() - 1 : 0;
}

function countRoster_() {
  try { return getRoster().length; } catch (e) { return 0; }
}
