/**
 * 担任・管理者向け — スプレッドシートのメニューと、Web管理画面の API。
 */

function onOpen() {
  buildMenu_();
}

/**
 * メニューは「ふだん使うもの」を上に、「一度きりの準備」を下にまとめる。
 * 項目が多いので、種類ごとにサブメニューへ入れている。
 */
function buildMenu_() {
  var ui = SpreadsheetApp.getUi();
  // 自動処理がオンかオフかは「設定」シートに書き写してある。1回だけ読む
  var flags = readAutoFlags_();

  ui.createMenu(MENU.ROOT)
    .addSubMenu(ui.createMenu(ngMenuLabel_(flags))
      .addItem('一覧シートを作る・更新する', 'menuRebuildNg')
      .addItem('指定を枠に反映する', 'menuApplyNg')
      .addSeparator()
      .addItem('チェックを自動で反映する（オン）', 'menuEnableNgAuto')
      .addItem('自動反映をやめる（オフ）', 'menuDisableNgAuto'))
    .addSubMenu(ui.createMenu(receptionMenuLabel_())
      .addItem('いまの受付状態を確認する', 'menuReceptionStatus')
      .addSeparator()
      .addItem('受付を開始する', 'menuPublish')
      .addItem('受付を停止する', 'menuUnpublish')
      .addSeparator()
      .addItem('保護者用URLを表示', 'menuShowUrl')
      .addSeparator()
      .addItem('未予約の生徒を担任にメールする', 'menuSendReminder')
      .addItem(autoLabel_(flags, REMINDER_AUTO_KEY, '締切前の自動リマインドを設定する'),
        'menuReminderSetting'))
    .addSeparator()
    .addItem(MENU.REFRESH, 'menuRefreshViews')
    .addSubMenu(ui.createMenu(MENU.PDF)
      .addItem('クラス別 予約表（全クラス）', 'menuExportAllPdf')
      .addItem('クラス別 予約表（開いているクラスのみ）', 'menuExportCurrentPdf')
      .addSeparator()
      .addItem('当日の面談記録シート（全クラス）', 'menuExportAllMeetingNotes')
      .addItem('当日の面談記録シート（開いているクラスのみ）', 'menuExportCurrentMeetingNotes')
      .addSeparator()
      .addSeparator()
      .addItem('当日の受付一覧（全校・時間順）', 'menuExportDaySchedule')
      .addSeparator()
      .addItem('保護者への案内プリント（QRコード付き）', 'menuExportHandout'))
    .addSeparator()
    .addSubMenu(ui.createMenu(MENU.ROWOPS)
      .addItem('面談を入れない（ブロック）にする', 'menuBlock')
      .addItem('ブロックを解除して空きに戻す', 'menuUnblock')
      .addSeparator()
      .addItem('この枠に生徒を入れる（代理登録）', 'menuAssignStudent')
      .addItem('予約を取り消す', 'menuCancel'))
    .addSeparator()
    .addSubMenu(ui.createMenu(MENU.MAINT)
      .addItem('① 初期セットアップ', 'menuSetup')
      .addItem('② 枠を再生成', 'menuGenerateSlots')
      .addItem('クラスの増減を反映する', 'menuSyncClasses')
      .addItem('シートの並び順を整える', 'menuOrderSheets')
      .addItem(autoLabel_(flags, VIEW_AUTO_KEY, '表示の自動更新を設定する'), 'menuAutoRefresh')
      .addSeparator()
      .addItem('🩺 データを点検する', 'menuCheckData')
      .addSeparator()
      .addItem('📦 バックアップを作成', 'menuBackup')
      .addItem(autoLabel_(flags, BACKUP_AUTO_KEY, '📦 毎日の自動バックアップを設定する'),
        'menuAutoBackup')
      .addSeparator()
      .addItem('③ ダミー生徒を作成（練習用・各クラス35名）', 'menuGenerateDummyRoster'))
    .addSeparator()
    .addSubMenu(ui.createMenu(MENU.RESET)
      .addItem('予約をすべて取り消す', 'menuClearBookings')
      .addItem('だめなコマの指定をすべて外す', 'menuClearNg')
      .addItem('予約ログを消す', 'menuClearLog')
      .addSeparator()
      .addItem('次の面談にそなえて初期化する', 'menuResetForNext')
      .addItem('年度末：すべて初期化する', 'menuResetAll')
      .addSeparator()
      .addItem('配布用にまっさらにする（他校へ渡す）', 'menuResetForDistribution'))
    .addToUi();
}

/**
 * 単体（非バインド）スクリプトの場合はこれを1度だけ手動実行して、
 * スプレッドシートを開いたときにメニューが出るようにする。
 */
function createOnOpenTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'onOpen') return 'すでに設定済みです。';
  }
  ScriptApp.newTrigger('onOpen').forSpreadsheet(ssId_()).onOpen().create();
  return 'メニューを表示するトリガーを設定しました。スプレッドシートを開き直してください。';
}

/* ---------------- メニュー処理 ---------------- */

function menuSetup() {
  var ui = SpreadsheetApp.getUi();
  var created = setupSystem();
  hideInternalSheets();
  ui.alert('初期セットアップ',
    (created.length ? '作成したシート: ' + created.join(', ') : '必要なシートはすべて揃っています。') +
    '\n\n次の順で入力してください。\n' +
    '1. 「クラス」シートに 4クラス分の担任名とメールを入力\n' +
    '2. 各「予約表_〇組」シートのA・B列にクラス・出席番号・生徒氏名を張り付け\n' +
    '3. 「面談日」「設定」を確認\n' +
    '4. メニューの「' + MENU.GENERATE + '」を実行\n' +
    '5. 「' + MENU.NG_SHEET + '」を実行し、面談を入れられない枠にチェック\n' +
    '　 → 「' + MENU.NG_APPLY + '」を実行\n' +
    '6. Webアプリをデプロイし、「' + MENU.PUBLISH + '」',
    ui.ButtonSet.OK);
}

function menuGenerateDummyRoster() {
  var ui = SpreadsheetApp.getUi();
  try {
    var res = generateDummyRoster();
    hideInternalSheets();
    var msg = '全4クラス分（計 ' + res.count + ' 名）のダミー生徒を各予約表シートのA・B列に追加しました。';
    if (res.slotsCreated) {
      msg += '\n\n※「枠マスタ」シートが空だったため、面談枠も自動的に再生成しました！';
    } else {
      msg += '\n\n※面談枠がまだない場合は、メニューの「' + MENU.GENERATE + '」を実行してください。';
    }
    ui.alert('ダミー生徒の作成', msg, ui.ButtonSet.OK);
  } catch (err) {
    ui.alert('エラー', String(err.message || err), ui.ButtonSet.OK);
  }
}

function menuGenerateSlots() {
  var ui = SpreadsheetApp.getUi();
  try {
    var r = generateSlots();
    rebuildOverview();
    rebuildClassSheets();
    hideInternalSheets();
    ui.alert('枠を再生成しました',
      r.written + ' 枠を作成しました（既存の予約 ' + r.kept + ' 件を引き継ぎ）。' +
      (r.ngBlocked ? '\n「' + SH.NG + '」シートの指定にしたがって ' + r.ngBlocked + ' 枠を面談なしにしました。' : '') +
      ngWarningText_(r.ngConflicts),
      ui.ButtonSet.OK);
  } catch (err) {
    ui.alert('再生成できません', String(err.message || err), ui.ButtonSet.OK);
  }
}

/**
 * クラスを増やした・減らした・名前を変えたあとの反映。
 * 使われなくなった名簿シートは自動で消さず、担当者に判断してもらう。
 */
function menuSyncClasses() {
  var ui = SpreadsheetApp.getUi();
  try {
    var res = syncClasses();

    var text = '面談枠を ' + res.slots + ' 枠にそろえました（既存の予約 ' + res.kept + ' 件を引き継ぎ）。';

    if (res.created.length) {
      text += '\n\n【作成した名簿シート】\n' + res.created.join('\n') +
        '\nこのシートのA・B列に、出席番号と生徒氏名を貼り付けてください。' +
        '\n貼り付けるまで、そのクラスの保護者は予約できません。';
    }

    if (res.orphans.length) {
      var lines = [];
      for (var i = 0; i < res.orphans.length; i++) {
        lines.push('・' + res.orphans[i].name + '（生徒 ' + res.orphans[i].students + ' 名分の名簿が残っています）');
      }
      text += '\n\n⚠【「' + SH.CLASSES + '」シートに無い名簿シート】\n' + lines.join('\n') +
        '\n\n・クラス名を変えた場合' + '\n' +
        '　上のシートのA・B列を、新しい名簿シートに貼り付けてください。' + '\n' +
        '・クラスを減らした場合' + '\n' +
        '　中身を確認のうえ、シートを手動で削除してください。個人情報が残ったままになります。' +
        '\n\n※誤って消さないよう、自動では削除していません。';
    }

    text += ngWarningText_(res.ngConflicts);

    ui.alert(res.orphans.length ? '⚠ 確認してください' : 'クラスの増減を反映しました',
      text, ui.ButtonSet.OK);
  } catch (err) {
    ui.alert('反映できません', String(err.message || err), ui.ButtonSet.OK);
  }
}

/** 入力ミスをまとめて洗い出す */
/* ---------------- 自動化の設定 ---------------- */

function menuEnableNgAuto() {
  var ui = SpreadsheetApp.getUi();
  try {
    var added = enableNgAutoApply();
    buildMenu_(); // サブメニュー名の「オン／手動」表示をその場で直す
    ui.alert(added ? '自動反映をオンにしました' : 'すでにオンになっています',
      '「' + SH.NG + '」シートのチェックを付け外しすると、その場で保護者の画面に反映されます。' + LF + LF +
      '「' + MENU.NG_APPLY + '」を実行する必要はなくなります（実行しても問題ありません）。' + LF +
      '予約が入っている枠は、これまでどおり変更されません。',
      ui.ButtonSet.OK);
  } catch (err) {
    ui.alert('設定できません',
      String(err.message || err) + LF + LF +
      '初回は承認画面が出ることがあります。表示に従って許可してから、もう一度実行してください。',
      ui.ButtonSet.OK);
  }
}

function menuDisableNgAuto() {
  var ui = SpreadsheetApp.getUi();
  try {
    var n = disableNgAutoApply();
    buildMenu_(); // サブメニュー名の「オン／手動」表示をその場で直す
    ui.alert(n ? '自動反映をオフにしました' : 'もともとオフです',
      'これからは「' + MENU.NG_APPLY + '」を実行したときだけ反映されます。', ui.ButtonSet.OK);
  } catch (err) {
    ui.alert('エラー', String(err.message || err), ui.ButtonSet.OK);
  }
}

function menuSendReminder() {
  var ui = SpreadsheetApp.getUi();
  try {
    var groups = unbookedByClass_();
    var lines = [], targets = 0, skip = [];

    for (var i = 0; i < groups.length; i++) {
      var g = groups[i];
      if (!g.students.length) { skip.push(g.cls + '：全員予約済み'); continue; }
      if (!g.email) { skip.push(g.cls + '：担任メールが未入力'); continue; }
      lines.push('・' + g.cls + '（未予約 ' + g.students.length + '名）→ ' + g.email);
      targets++;
    }

    if (!targets) {
      ui.alert('送信できるクラスがありません',
        (skip.length ? skip.join(LF) : '未予約の生徒がいません。'), ui.ButtonSet.OK);
      return;
    }

    var res = ui.alert('未予約の生徒一覧を担任へ送ります',
      '次のあて先にメールを送信します。よろしいですか？' + LF + LF +
      lines.join(LF) +
      (skip.length ? LF + LF + '【送信しないクラス】' + LF + skip.join(LF) : ''),
      ui.ButtonSet.OK_CANCEL);
    if (res !== ui.Button.OK) return;

    var sent = sendUnbookedReminder();
    ui.alert('送信しました',
      sent.sent.length + ' クラスの担任へ送信しました。' +
      (sent.done.length ? LF + '全員予約済みのクラス: ' + sent.done.join('、') : '') +
      (sent.noEmail.length ? LF + '送信できなかったクラス: ' + sent.noEmail.join('、') : ''),
      ui.ButtonSet.OK);
  } catch (err) {
    ui.alert('送信できません', String(err.message || err), ui.ButtonSet.OK);
  }
}

function menuReminderSetting() {
  var ui = SpreadsheetApp.getUi();
  try {
    var cfg = getConfig();
    var on = reminderEnabled_();

    if (on) {
      var off = ui.alert('自動リマインドは現在オンです',
        '締切の ' + cfg.reminderDays + ' 日前の朝に、未予約の生徒一覧を担任へ自動送信します。' + LF + LF +
        'オフにしますか？', ui.ButtonSet.OK_CANCEL);
      if (off !== ui.Button.OK) return;
      disableReminder();
      buildMenu_();
      ui.alert('自動リマインドをオフにしました。', ui.ButtonSet.OK);
      return;
    }

    if (!cfg.closeAt) {
      ui.alert('先に締切を設定してください',
        '「' + SH.CONFIG + '」シートの「予約受付締切」に日時を入れてから、もう一度実行してください。' + LF +
        '締切を基準に送信日を決めるためです。', ui.ButtonSet.OK);
      return;
    }

    var target = new Date(cfg.closeAt.getTime());
    target.setDate(target.getDate() - cfg.reminderDays);

    var ok = ui.alert('締切前の自動リマインドをオンにしますか？',
      '送信予定日: ' + Utilities.formatDate(target, TZ, 'M月d日') + ' の朝' + LF +
      '（締切 ' + Utilities.formatDate(cfg.closeAt, TZ, 'M月d日 HH:mm') + ' の ' + cfg.reminderDays + ' 日前）' + LF + LF +
      'その日に未予約の生徒がいるクラスの担任へ、一覧をメールで送ります。' + LF +
      '日数は「' + SH.CONFIG + '」シートの「リマインド日数(締切の何日前)」で変えられます。',
      ui.ButtonSet.OK_CANCEL);
    if (ok !== ui.Button.OK) return;

    enableReminder();
    buildMenu_();
    ui.alert('自動リマインドをオンにしました',
      '毎朝チェックし、該当の日だけ送信します。', ui.ButtonSet.OK);
  } catch (err) {
    ui.alert('設定できません', String(err.message || err), ui.ButtonSet.OK);
  }
}

/**
 * 受付期間中の自動バックアップ。
 * 予約が消えると復旧できないので、いちばん失いたくない時期だけ自動で守る。
 */
function menuAutoBackup() {
  var ui = SpreadsheetApp.getUi();
  try {
    if (autoBackupEnabled_()) {
      var off = ui.alert('毎日の自動バックアップは現在オンです',
        '受付が開いている間、毎日夜中にバックアップを取っています。' + LF +
        '受付を停止しているときや、締切を過ぎたあとは動きません。' + LF + LF +
        'オフにしますか？', ui.ButtonSet.OK_CANCEL);
      if (off !== ui.Button.OK) return;
      disableAutoBackup();
      buildMenu_();
      ui.alert('自動バックアップをオフにしました。', ui.ButtonSet.OK);
      return;
    }

    var ok = ui.alert('毎日の自動バックアップをオンにしますか？',
      '受付が開いている間、毎日夜中に「' + BACKUP_FOLDER_NAME + '」の中の' + LF +
      '「' + AUTO_BACKUP_FOLDER_NAME + '」へ複製を保存します。' + LF +
      '手で取ったバックアップとは別のフォルダなので、混ざりません。' + LF + LF +
      '【なぜ必要か】' + LF +
      '予約が消えたり名簿を上書きしてしまうと、元には戻せません。' + LF +
      '手で取るのを忘れた日があると、その日の分は取り返しがつきません。' + LF + LF +
      '受付を停止したあとや締切を過ぎたあとは動かないので、' + LF +
      'ファイルが際限なく増えることはありません。' + LF + LF +
      '古い自動バックアップは、新しいものから30個を残して' + LF +
      'Drive のゴミ箱へ移します（30日間は戻せます）。' + LF +
      '残す数は「' + SH.CONFIG + '」シートの「' + AUTO_BACKUP_KEEP_KEY + '」で変えられます。',
      ui.ButtonSet.OK_CANCEL);
    if (ok !== ui.Button.OK) return;

    enableAutoBackup();
    buildMenu_();
    ensureConfigKey_(AUTO_BACKUP_KEEP_KEY,
      '自動バックアップを何個まで残すか。これを超えた古いものはゴミ箱へ移す。0で無制限', 30);
    ui.alert('自動バックアップをオンにしました',
      '受付中は毎日夜中に保存します。' + LF +
      '保存先は「' + BACKUP_FOLDER_NAME + ' ▸ ' + AUTO_BACKUP_FOLDER_NAME + '」です。' + LF +
      '残す世代数は「' + SH.CONFIG + '」シートの「' + AUTO_BACKUP_KEEP_KEY + '」で調整できます。' + LF + LF +
      'いますぐ取りたいときは「' + MENU.BACKUP + '」を実行してください。', ui.ButtonSet.OK);
  } catch (err) {
    ui.alert('設定できません', String(err.message || err), ui.ButtonSet.OK);
  }
}

function menuCheckData() {
  var ui = SpreadsheetApp.getUi();
  ss_().toast('データを点検しています…', '三者面談 点検', 10);
  try {
    var res = checkData();
    var title = res.errors.length ? '🩺 ' + res.errors.length + ' 件、直したほうがよい点があります'
      : (res.warns.length ? '🩺 ' + res.warns.length + ' 件、確認したい点があります'
        : '🩺 点検が終わりました');
    ui.alert(title, formatCheckResult_(res), ui.ButtonSet.OK);
  } catch (err) {
    ui.alert('点検できません', String(err.message || err), ui.ButtonSet.OK);
  }
}

/**
 * 表（全体ビュー・クラス別予約表）の自動更新。
 * 予約のたびに作り直すと全校規模では追いつかないため、
 * まとめて数分おきに反映する。保護者の予約可否には影響しない。
 */
function menuAutoRefresh() {
  var ui = SpreadsheetApp.getUi();
  try {
    if (autoRefreshEnabled_()) {
      var off = ui.alert('表示の自動更新は現在オンです',
        '数分おきに、たまった予約を全体ビューとクラス別予約表へ反映しています。' + LF + LF +
        'オフにしますか？（オフにすると「' + MENU.REFRESH + '」を押したときだけ反映されます）',
        ui.ButtonSet.OK_CANCEL);
      if (off !== ui.Button.OK) return;
      disableAutoRefresh();
      buildMenu_();
      ui.alert('自動更新をオフにしました。', ui.ButtonSet.OK);
      return;
    }

    var ok = ui.alert('表示の自動更新をオンにしますか？',
      '5分おきに、たまった予約を表へ反映します。' + LF + LF +
      '【なぜ必要か】' + LF +
      '予約のたびに全クラスの表を作り直すと、クラス数が多い学校では' + LF +
      '1件あたり1分以上かかり、保護者の待ち行列になってしまいます。' + LF +
      'まとめて反映することで、予約は数秒で確定するようになります。' + LF + LF +
      '表示は最大5分遅れますが、保護者が予約できる・できないの判定は' + LF +
      '常に最新です。二重予約は起きません。',
      ui.ButtonSet.OK_CANCEL);
    if (ok !== ui.Button.OK) return;

    var m = enableAutoRefresh(5);
    buildMenu_();
    ui.alert('自動更新をオンにしました', m + '分おきに反映します。' + LF +
      'すぐ見たいときは「' + MENU.REFRESH + '」を押してください。', ui.ButtonSet.OK);
  } catch (err) {
    ui.alert('設定できません',
      String(err.message || err) + LF + LF +
      '初回は承認画面が出ることがあります。表示に従って許可してから、もう一度実行してください。',
      ui.ButtonSet.OK);
  }
}

function menuOrderSheets() {
  var ui = SpreadsheetApp.getUi();
  try {
    var n = orderSheets();
    var rows = orderConfigRows();
    ui.alert('並び順を整えました',
      'シート ' + n + ' 枚を左から「設定・面談日・クラス・' + SH.NG + '・' + SH.LOG + '・' + SH.OVERVIEW +
      '・予約表（クラス順）」の順に並べました。' + LF +
      'この一覧に無いシートは動かしていません。' + LF + LF +
      '「' + SH.CONFIG + '」シートの ' + rows + ' 行を、意味のまとまりごとに並べ直しました。' + LF +
      '（対で使う「リマインド日数」と「締切前の自動リマインド」などが隣り合います）',
      ui.ButtonSet.OK);
  } catch (err) {
    ui.alert('エラー', String(err.message || err), ui.ButtonSet.OK);
  }
}

function menuRefreshViews() {
  var pending = pendingViewUpdates_();
  refreshViews(true);

  // よく押される項目なので、ここで自動処理の表示を実態に合わせ直しておく。
  // 「設定」シートの表示用の行を手で書き換えられていても、これで元に戻る。
  try {
    // 担任が手で直した内容を確実に反映させる
    dropRefCaches_();
    ensureConfigWatcher();
    syncAdminUrl_();
    // 古いキー名が残っていれば付け替える。名前を直したあと、
    // 並び順を整える機会が無いまま使い続けられることがあるため、ここでも見る
    migrateConfigKeys_();
    // あとから増えた設定は、初期セットアップを実行しないと行ができない。
    // よく押される項目なので、ここで用意しておく
    ensureConfigKey_(RESERVE_COUNT_KEY,
      '保護者には見せない予備の枠。最終コマの後ろに作られ、担任が管理画面・行メニュー・黄色い行のL〜N列への直接入力で埋める。0で無し', 0);
    if (syncAllAutoFlags_()) buildMenu_();
  } catch (e) {
    console.warn('自動処理の状態の同期をスキップ:', e);
  }
  ss_().toast(
    pending ? pending + ' 件の予約を反映しました。' : '最新の状態です。',
    '三者面談 表示の更新', 5);
}

function menuExportAllPdf() {
  var ui = SpreadsheetApp.getUi();
  ss_().toast('全クラスの予約表PDFを作成しています。十数秒お待ちください…', '三者面談 PDF出力', 15);
  try {
    var res = exportAllClassesPdf();
    showAllPdfCompleteDialog_(res, false);
  } catch (err) {
    ui.alert('PDF作成エラー', String(err.message || err), ui.ButtonSet.OK);
  }
}

function menuExportCurrentPdf() {
  var ui = SpreadsheetApp.getUi();
  ss_().toast('予約表PDFを作成しています。少々お待ちください…', '三者面談 PDF出力', 10);
  try {
    var res = exportCurrentClassPdf();
    showPdfCompleteDialog_(res, false);
  } catch (err) {
    ui.alert('PDF作成エラー', String(err.message || err), ui.ButtonSet.OK);
  }
}

/** 当日の受付本部用（全校・時間順） */
function menuExportDaySchedule() {
  var ui = SpreadsheetApp.getUi();
  ss_().toast('当日受付一覧を作成しています…', '三者面談 PDF', 15);
  try {
    var res = exportDaySchedulePdf();
    ui.alert('当日受付一覧を作成しました',
      res.fileName + LF + LF +
      '来校する ' + res.count + ' 件を、日付 → 時間 → クラスの順に並べています。' + LF +
      '受付に1枚置いておくと、来校の確認に使えます。' + LF + LF +
      '保存先: ' + res.folderName, ui.ButtonSet.OK);
  } catch (err) {
    ui.alert('作成できません', String(err.message || err), ui.ButtonSet.OK);
  }
}

/** 保護者に配る案内プリント（A4・QRコード付き） */
function menuExportHandout() {
  var ui = SpreadsheetApp.getUi();
  ss_().toast('案内プリントを作成しています…', '三者面談 配布プリント', 15);
  try {
    var res = exportHandoutPdf();
    showHandoutCompleteDialog_(res);
  } catch (err) {
    ui.alert('作成できません', String(err.message || err), ui.ButtonSet.OK);
  }
}

function showHandoutCompleteDialog_(res) {
  var html = '<!DOCTYPE html><html><head><base target="_blank">' +
    '<style>' +
    'body { font-family: -apple-system, "Segoe UI", Roboto, sans-serif; padding:16px 20px; color:#202124; line-height:1.6; margin:0; }' +
    'h2 { font-size:1.12rem; margin:0 0 10px; color:#1a73e8; }' +
    'p { margin:0 0 10px; font-size:.94rem; }' +
    '.u { font-size:.8rem; color:#5f6368; background:#f1f3f4; padding:6px 10px; border-radius:6px; word-break:break-all; margin:8px 0 14px; }' +
    'a.btn { display:block; text-align:center; padding:10px 14px; text-decoration:none; border-radius:6px; font-weight:600; font-size:.94rem; margin-bottom:9px; }' +
    'a.p { background:#1a73e8; color:#fff; }' +
    'a.s { background:#fff; color:#1a73e8; border:1px solid #dadce0; }' +
    'button { width:100%; padding:8px; background:transparent; border:none; color:#5f6368; font-size:.87rem; cursor:pointer; }' +
    '</style></head><body>' +
    '<h2>案内プリントを作成しました</h2>' +
    '<p>A4・1枚です。印刷して配布するか、PDFのまま配信してください。</p>' +
    '<p style="font-size:.86rem;color:#5f6368">QRコードは、下のURLを読み取るように作られています。' +
    '<strong>配布前に、ご自身のスマートフォンで一度読み取って確認してください。</strong></p>' +
    '<div class="u">' + escHtml_(res.url) + '</div>' +
    '<a class="btn p" href="' + escHtml_(res.fileUrl) + '">📄 プリントを開く</a>' +
    '<a class="btn s" href="' + escHtml_(res.folderUrl) + '">📁 ' + escHtml_(res.folderName) + ' を開く</a>' +
    '<button onclick="google.script.host.close()">閉じる</button>' +
    '</body></html>';

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(420).setHeight(330),
    '配布プリント');
}

function menuExportAllMeetingNotes() {
  var ui = SpreadsheetApp.getUi();
  ss_().toast('全クラスの当日面談記録シートを作成しています。十数秒お待ちください…', '三者面談 面談シート出力', 15);
  try {
    var res = exportAllMeetingNotesPdf();
    showAllPdfCompleteDialog_(res, true);
  } catch (err) {
    ui.alert('面談シート作成エラー', String(err.message || err), ui.ButtonSet.OK);
  }
}

function menuExportCurrentMeetingNotes() {
  var ui = SpreadsheetApp.getUi();
  ss_().toast('当日面談記録シートを作成しています。少々お待ちください…', '三者面談 面談シート出力', 10);
  try {
    var res = exportCurrentMeetingNotesPdf();
    showPdfCompleteDialog_(res, true);
  } catch (err) {
    ui.alert('面談シート作成エラー', String(err.message || err), ui.ButtonSet.OK);
  }
}

function menuBackup() {
  var ui = SpreadsheetApp.getUi();
  ss_().toast('バックアップを作成しています…', '三者面談 バックアップ', 5);
  try {
    var res = createManualBackup();
    showBackupCompleteDialog_(res);
  } catch (err) {
    ui.alert('バックアップエラー', String(err.message || err), ui.ButtonSet.OK);
  }
}

/* ---------------- だめなコマ（担任が面談を入れない枠） ---------------- */

function menuRebuildNg() {
  var ui = SpreadsheetApp.getUi();
  try {
    var n = rebuildNgSheet();
    if (!n) {
      ui.alert('「面談日」で「実施する」がTRUEの日付と、「クラス」シートのクラス名を先に入力してください。');
      return;
    }
    var sh = ss_().getSheetByName(SH.NG);
    if (sh) ss_().setActiveSheet(sh);
    ui.alert('だめなコマシートを更新しました',
      '「' + SH.NG + '」シートに ' + n + ' 行（面談日 × コマ）を用意しました。\n\n' +
      '面談を入れられない枠に、自分のクラスの列でチェックを入れてください。\n' +
      '入力が終わったら、メニューの「' + MENU.NG_APPLY + '」を実行します。\n' +
      '（すでに入っているチェックはそのまま引き継いでいます）',
      ui.ButtonSet.OK);
  } catch (err) {
    ui.alert('エラー', String(err.message || err), ui.ButtonSet.OK);
  }
}

function menuApplyNg() {
  var ui = SpreadsheetApp.getUi();
  try {
    var res = applyNgSlots();
    rebuildOverview();
    rebuildClassSheets();
    hideInternalSheets();

    var text;
    if (!res.blocked && !res.unblocked) {
      text = 'すでに「' + SH.NG + '」シートのとおりに反映されています。';
    } else {
      text = 'チェックした ' + res.blocked + ' 枠を「面談なし」にしました。';
      if (res.unblocked) text += '\nチェックを外した ' + res.unblocked + ' 枠を「空き」に戻しました。';
    }
    text += ngWarningText_(res.conflicts);
    ui.alert(res.conflicts.length ? '⚠ 確認してください' : 'だめなコマを反映しました',
      text, ui.ButtonSet.OK);
  } catch (err) {
    ui.alert('反映できません', String(err.message || err), ui.ButtonSet.OK);
  }
}

/**
 * 「だめなコマ」に指定したのに予約が残っている枠についての注意喚起文。
 * 対象が無ければ空文字を返す。
 */
function ngWarningText_(conflicts) {
  if (!conflicts || !conflicts.length) return '';
  var lines = [];
  for (var i = 0; i < conflicts.length; i++) {
    lines.push('・' + (typeof conflicts[i] === 'string' ? conflicts[i] : ngConflictLabel_(conflicts[i])));
  }
  return '\n\n⚠ 次の ' + conflicts.length + ' 件は、すでに保護者の予約が入っているため面談なしにできませんでした。\n' +
    lines.join('\n') + '\n\n' +
    'この時間には面談の予約が残ったままです。\n' +
    '保護者に連絡して面談時間を移すか、予約を取り消してください。\n' +
    'そのあと「' + MENU.NG_APPLY + '」をもう一度実行すると、警告が消えます。\n' +
    '（該当のセルは「' + SH.NG + '」シートで赤く表示され、メモに生徒名が入っています）';
}

function menuBlock() { setStatusForSelection_(STATUS.BLOCKED); }
function menuUnblock() { setStatusForSelection_(STATUS.OPEN); }

/**
 * 選択範囲の行から枠IDを集める。
 * 「枠マスタ」と各「予約表_〇組」シートに対応。対応外のシートなら null を返す。
 * @return {Array<string>|null}
 */
function selectedSlotIds_(sh, start, end) {
  var name = sh.getName();
  var count = end - start + 1;
  if (count < 1) return [];

  if (name === SH.SLOTS) {
    var vals = sh.getRange(start, COL.SLOT_ID, count, 1).getValues();
    var ids = [];
    for (var i = 0; i < vals.length; i++) {
      if (vals[i][0]) ids.push(String(vals[i][0]));
    }
    return ids;
  }

  if (name.indexOf(CLASS_SHEET_PREFIX) === 0) {
    var cls = name.slice(CLASS_SHEET_PREFIX.length);
    var index = {};
    var slots = readSlots_();
    for (var s = 0; s < slots.length; s++) {
      var v = slots[s].v;
      if (String(v[COL.CLASS - 1]) !== cls) continue;
      index[dateLabel_(v[COL.DATE - 1]) + '|' + v[COL.START - 1] + '–' + v[COL.END - 1]] =
        String(v[COL.SLOT_ID - 1]);
    }
    // 右側の「時間枠別 予約表」は I列=日付、J列=時間
    var right = sh.getRange(start, 9, count, 2).getValues();
    var out = [];
    for (var r = 0; r < right.length; r++) {
      var key = String(right[r][0] || '').trim() + '|' + String(right[r][1] || '').trim();
      if (index[key]) out.push(index[key]);
    }
    return out;
  }

  return null;
}

/**
 * 枠IDをまとめて指定の状態にし、「だめなコマ」シートのチェックも同期する。
 * 予約が入っている枠は変更しない。
 * @return {{changed:number, skipped:number}}
 */
function setStatusForSlotIds_(slotIds, newStatus) {
  return withLock_(function () {
    var sh = sheet_(SH.SLOTS);
    var last = sh.getLastRow();
    if (last < 2) return { changed: 0, skipped: 0 };

    var slots = readSlots_();
    var byId = {};
    for (var i = 0; i < slots.length; i++) byId[String(slots[i].v[COL.SLOT_ID - 1])] = slots[i];

    var statusCol = sh.getRange(2, COL.STATUS, last - 1, 1).getValues();
    var changed = 0, skipped = 0, touched = [], seen = {}, skippedList = [];

    for (var k = 0; k < slotIds.length; k++) {
      var id = String(slotIds[k]);
      if (seen[id]) continue;
      seen[id] = true;
      var found = byId[id];
      if (!found) continue;
      var cur = String(found.v[COL.STATUS - 1]);
      if (cur === STATUS.BOOKED) {
        skipped++;
        skippedList.push({
          slotId: id,
          dateLabel: dateLabel_(found.v[COL.DATE - 1]),
          time: String(found.v[COL.START - 1]) + '–' + String(found.v[COL.END - 1]),
          cls: String(found.v[COL.CLASS - 1]),
          no: found.v[COL.NUMBER - 1],
          name: String(found.v[COL.STUDENT - 1] || ''),
          guardian: String(found.v[COL.GUARDIAN - 1] || '')
        });
        continue;
      }
      if (cur === newStatus) continue;
      statusCol[found.row - 2][0] = newStatus;
      touched.push({ slotId: id, flag: newStatus === STATUS.BLOCKED });
      changed++;
    }

    if (changed) {
      sh.getRange(2, COL.STATUS, last - 1, 1).setValues(statusCol);
      clearSlotCache_();
      markViewsStale_();
      try {
        setNgFlags_(touched);
      } catch (e) {
        console.warn('だめなコマシートの同期をスキップ:', e);
      }
    }
    return { changed: changed, skipped: skipped, skippedList: skippedList };
  });
}

function setStatusForSelection_(newStatus) {
  var ui = SpreadsheetApp.getUi();
  var sh = SpreadsheetApp.getActiveSheet();

  if (sh.getName() === SH.NG) {
    ui.alert('「' + SH.NG + '」シートでは、クラスの列にチェックを入れてから\n' +
      'メニューの「' + MENU.NG_APPLY + '」を実行してください。');
    return;
  }

  var rng = SpreadsheetApp.getActiveRange();
  var start = Math.max(rng.getRow(), 2);
  var end = rng.getRow() + rng.getNumRows() - 1;
  if (end < start) { ui.alert('見出し行以外の行を選んでください。'); return; }

  var ids = selectedSlotIds_(sh, start, end);
  if (ids === null) {
    ui.alert('「' + SH.SLOTS + '」シートまたは各「予約表_〇組」シートで、対象の行を選んでから実行してください。\n' +
      'まとめて指定するときは「' + SH.NG + '」シートが便利です。');
    return;
  }
  if (!ids.length) { ui.alert('選んだ範囲に面談枠がありません。'); return; }

  var res = setStatusForSlotIds_(ids, newStatus);
  rebuildOverview();
  rebuildClassSheets();
  hideInternalSheets();
  ui.alert(res.skipped ? '⚠ 確認してください' : '完了',
    res.changed + ' 件を「' + newStatus + '」にしました。' + ngWarningText_(res.skippedList),
    ui.ButtonSet.OK);
}

/**
 * 選んだ枠に、担任が代理で生徒を入れる。
 *
 * 「予約表_〇組」の右側の表で行を選んで実行する。
 * この表は自動で作り直されるため直接書いても残らないが、
 * 電話で申し込まれた場合など、シートを見ながら登録したい場面はある。
 */
function menuAssignStudent() {
  var ui = SpreadsheetApp.getUi();
  try {
    var slotId = selectedSlotId_();
    if (!slotId) {
      ui.alert('枠が選ばれていません',
        '「' + CLASS_SHEET_PREFIX + '〇組」シートの右側の表で、' + LF +
        '入れたい時間の行をどこか1つ選んでから実行してください。',
        ui.ButtonSet.OK);
      return;
    }

    var t = HtmlService.createTemplateFromFile('assign');
    t.slotId = slotId;
    ui.showModalDialog(t.evaluate().setWidth(380).setHeight(340), 'この枠に生徒を入れる');
  } catch (err) {
    ui.alert('開けません', String(err.message || err), ui.ButtonSet.OK);
  }
}

/** いま選ばれている行の枠ID。枠マスタでも予約表でも拾える */
function selectedSlotId_() {
  var sh = SpreadsheetApp.getActiveSheet();
  var name = sh.getName();
  var row = SpreadsheetApp.getActiveRange().getRow();
  var id = '';

  if (name === SH.SLOTS) {
    id = String(sh.getRange(row, COL.SLOT_ID).getValue() || '').trim();
  } else if (name.indexOf(CLASS_SHEET_PREFIX) === 0) {
    // 右側の表の末尾に、見えない状態で枠IDを持たせてある
    id = String(sh.getRange(row, 9 + CLASS_HEADER_RIGHT.length).getValue() || '').trim();
  }
  return (!id || id === '枠ID') ? '' : id;
}

/** 小窓に渡す情報。枠の日時と、そのクラスの名簿 */
function dialogAssignInfo(slotId) {
  return safe_(function () {
    var slot = findSlotRow_(slotId);
    var cls = String(slot.v[COL.CLASS - 1]);
    var roster = getRoster();
    var students = [];
    for (var i = 0; i < roster.length; i++) {
      if (roster[i].cls === cls) students.push({ no: roster[i].no, name: roster[i].name });
    }
    return {
      cls: cls,
      when: dateLabel_(slot.v[COL.DATE - 1]) + ' ' +
        slot.v[COL.START - 1] + '–' + slot.v[COL.END - 1],
      // 予備は保護者に見えない枠なので、予約コードを出さない。小窓の文面もそれに合わせる
      reserve: String(slot.v[COL.STATUS - 1]) === STATUS.RESERVE,
      // 予備の行はシートへ直接書けないので、入れ直し・取り消しもこの小窓で行う
      currentNo: slot.v[COL.NUMBER - 1] || '',
      currentName: String(slot.v[COL.STUDENT - 1] || ''),
      students: students
    };
  });
}

/**
 * 小窓からの登録。終わったら表も作り直す。
 *
 * 予備の枠は「予約」ではなく、担任が使う枠への記入として扱う。
 * Web管理画面が黄色いセルを押したときと同じ結果になる。
 */
function dialogAssign(slotId, text) {
  return safe_(function () {
    var reserve = String(findSlotRow_(slotId).v[COL.STATUS - 1]) === STATUS.RESERVE;
    var done = reserve
      ? setReserveStudent_(slotId, text, '担任シート')
      : assignStudentToSlot_(slotId, text);
    done.reserve = reserve;
    try {
      refreshAfterReset_();
    } catch (e) {
      console.warn('表の更新をスキップ:', e);
    }
    return done;
  });
}

function menuCancel() {
  var ui = SpreadsheetApp.getUi();
  var sh = SpreadsheetApp.getActiveSheet();
  if (sh.getName() !== SH.SLOTS) {
    ui.alert('「' + SH.SLOTS + '」シートで、取り消したい予約の行を選んでから実行してください。');
    return;
  }
  var rng = SpreadsheetApp.getActiveRange();
  var start = Math.max(rng.getRow(), 2);
  var end = rng.getRow() + rng.getNumRows() - 1;

  var targets = [];
  for (var r = start; r <= end; r++) {
    var v = sh.getRange(r, 1, 1, SLOT_LAST_COL).getValues()[0];
    if (String(v[COL.STATUS - 1]) === STATUS.BOOKED) targets.push({ row: r, v: v });
  }
  if (!targets.length) { ui.alert('選んだ範囲に予約はありません。'); return; }

  var names = targets.map(function (t) {
    return dateLabel_(t.v[COL.DATE - 1]) + ' ' + t.v[COL.START - 1] + '　' +
      t.v[COL.CLASS - 1] + ' ' + t.v[COL.NUMBER - 1] + '. ' + t.v[COL.STUDENT - 1];
  }).join('\n');
  var res = ui.alert('次の予約を取り消します。よろしいですか？', names, ui.ButtonSet.OK_CANCEL);
  if (res !== ui.Button.OK) return;

  targets.forEach(function (t) {
    clearSlotRow_(t.row, String(t.v[COL.SLOT_ID - 1]));
    logAction_('取消', String(t.v[COL.SLOT_ID - 1]), String(t.v[COL.CLASS - 1]),
      t.v[COL.NUMBER - 1], String(t.v[COL.STUDENT - 1]),
      '担任による取消 / コード ' + String(t.v[COL.CODE - 1] || ''));
  });
  clearSlotCache_();
  rebuildOverview();
  rebuildClassSheets();
  hideInternalSheets();
  ui.alert(targets.length + ' 件の予約を取り消しました。');
}

/**
 * 受付を開始する。
 * 開けてしまうと保護者が押し寄せるので、直せるうちに問題を出しきる。
 */
function menuPublish() {
  var ui = SpreadsheetApp.getUi();
  ss_().toast('受付開始前の点検をしています…', '三者面談', 15);

  var report;
  try {
    report = publishPreflight_();
  } catch (e) {
    report = { blockers: [], warnings: ['点検を実行できませんでした: ' + (e.message || e)] };
  }

  if (report.blockers.length) {
    var ok = ui.alert('⚠ このままでは保護者が予約できません',
      report.blockers.join(LF + LF) + LF + LF +
      '直してから開始することをおすすめします。' + LF +
      'それでも受付を開始しますか？',
      ui.ButtonSet.OK_CANCEL);
    if (ok !== ui.Button.OK) return;
  } else if (report.warnings.length) {
    var ok2 = ui.alert('確認したい点があります',
      report.warnings.join(LF + LF) + LF + LF + '受付を開始しますか？',
      ui.ButtonSet.OK_CANCEL);
    if (ok2 !== ui.Button.OK) return;
  }

  setPublished_(true);
}

/**
 * 受付を開始してよいかを調べる。
 * データの点検に加えて、保護者用URLに実際にアクセスできるかまで確かめる。
 * @return {{blockers:Array<string>, warnings:Array<string>}}
 */
function publishPreflight_() {
  var blockers = [], warnings = [];

  try {
    var res = checkData();
    for (var i = 0; i < res.errors.length; i++) {
      blockers.push(formatFinding_('🔴', res.errors[i]));
    }
    for (var w = 0; w < res.warns.length; w++) {
      warnings.push(formatFinding_('🟡', res.warns[w]));
    }
  } catch (e) {
    warnings.push('🟡 データの点検を実行できませんでした: ' + (e.message || e));
  }

  var reach = parentUrlReachability_();
  if (reach.level === 'error') blockers.push('🔴 ' + reach.message);
  else if (reach.level === 'warn') warnings.push('🟡 ' + reach.message);

  return { blockers: blockers, warnings: warnings };
}

function formatFinding_(mark, f) {
  return mark + ' ' + f.title +
    (f.detail ? '（' + f.detail + '）' : '') +
    (f.fix ? LF + '　→ ' + f.fix : '');
}

/**
 * 保護者用URLに、ログインしていない状態でたどり着けるか。
 *
 * UrlFetchApp はブラウザのログイン情報を持たないので、
 * 保護者とほぼ同じ条件で試せる。デプロイの設定間違いはここで出る。
 */
/** 転送先のURL（ヘッダの大文字小文字はGoogle側で変わることがある） */
function redirectLocation_(res) {
  try {
    var h = res.getAllHeaders() || {};
    return String(h['Location'] || h['location'] || '');
  } catch (e) {
    return '';
  }
}

function loginRequiredMessage_(code) {
  return '保護者用URLを開くとログインを求められます（応答 ' + code + '）。' + LF +
    '　この状態では、保護者はページを開けません。' + LF +
    '　→ Apps Script エディタの「デプロイ ▸ デプロイを管理 ▸ 編集」で' + LF +
    '　　「次のユーザーとして実行: 自分」「アクセスできるユーザー: 全員」に直してください。';
}

function parentUrlReachability_() {
  var url = webAppUrl_();
  if (!url) {
    return {
      level: 'error',
      message: '保護者用URLがまだありません。' + LF +
        '　→ ウェブアプリを公開し、そのURLを「' + SH.CONFIG + '」シートの「保護者用URL」に貼り付けてください。'
    };
  }

  try {
    // 画面のHTMLが200でも、Webアプリから設定シートを開けないことがある。
    // health=1 は getConfig() まで実行するので、保護者が予約操作できる状態を確認できる。
    var healthUrl = url + (url.indexOf('?') >= 0 ? '&' : '?') +
      'health=1&cb=' + new Date().getTime();
    var res = UrlFetchApp.fetch(healthUrl, { muteHttpExceptions: true, followRedirects: true });
    var code = res.getResponseCode();

    if (code === 200) {
      var body = String(res.getContentText() || '').trim();
      try {
        var health = JSON.parse(body);
        if (health && health.ok) return { level: 'ok', message: '' };
        if (health && health.error) {
          return {
            level: 'error',
            message: '予約ページは表示できますが、予約データを開けません。' + LF +
              '　' + health.error + LF +
              '　→ スプレッドシートを再読み込みしてから、Webアプリのデプロイを更新してください。'
          };
        }
      } catch (parseError) {
        // health=1 に対応していない旧デプロイは、通常のHTMLを返す。
      }
      return {
        level: 'error',
        message: '予約ページが古いデプロイのままです。' + LF +
          '　→ Apps Script エディタの「デプロイ ▸ デプロイを管理 ▸ 編集」で' + LF +
          '　　バージョンを「新バージョン」にしてデプロイを更新してください。'
      };
    }

    // 正しく公開できていても、いったん googleusercontent へ転送されることがある。
    // 転送先を見ないと、動いているものまでエラー扱いしてしまう。
    if (code === 301 || code === 302 || code === 303 || code === 307) {
      var loc = redirectLocation_(res);
      if (loc.indexOf('googleusercontent.com') >= 0) return { level: 'ok', message: '' };
      if (loc.indexOf('accounts.google.com') >= 0 || loc.indexOf('ServiceLogin') >= 0) {
        return { level: 'error', message: loginRequiredMessage_(code) };
      }
      return {
        level: 'warn',
        message: '保護者用URLが別の場所へ転送されました' +
          (loc ? '（' + loc.split('?')[0] + '）' : '') + '。ブラウザで開けるか確かめてください。'
      };
    }

    if (code === 401 || code === 403) {
      return { level: 'error', message: loginRequiredMessage_(code) };
    }

    if (code === 404) {
      return {
        level: 'error',
        message: '保護者用URLが見つかりません（404）。' + LF +
          '　デプロイが削除されたか、アーカイブされている可能性があります。' + LF +
          '　→ 「デプロイ ▸ デプロイを管理」で生きているURLを確かめ、' + LF +
          '　　「' + SH.CONFIG + '」シートの「保護者用URL」に貼り直してください。'
      };
    }

    return {
      level: 'warn',
      message: '保護者用URLの応答が ' + code + ' でした。ブラウザで開けるか確かめてください。'
    };
  } catch (e) {
    return {
      level: 'warn',
      message: '保護者用URLに接続できませんでした（' + (e.message || e) + '）。'
    };
  }
}
function menuUnpublish() { setPublished_(false); }

function setPublished_(flag) {
  var sh = sheet_(SH.CONFIG);
  var vals = sh.getDataRange().getValues();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === '公開') {
      sh.getRange(i + 1, 2).setValue(flag);
      // サブメニューに出している「受付中／停止中」の表示を、その場で直す
      try { buildMenu_(); } catch (e) { /* メニューが作れなくても受付の切替は済んでいる */ }
      SpreadsheetApp.getUi().alert('予約受付を' + (flag ? '開始' : '停止') + 'しました。');
      return;
    }
  }
  SpreadsheetApp.getUi().alert('「設定」シートに「公開」の行がありません。「' + MENU.SETUP + '」を実行してください。');
}

function menuShowUrl() {
  var ui = SpreadsheetApp.getUi();

  // 貼り付け先が無いと直しようがないので、行だけ先に用意しておく
  var added = false;
  try {
    added = ensureConfigKey_('保護者用URL',
      '実際に保護者へ配るURL(末尾 /exec)。空欄なら自動取得するが、デプロイが複数あると別のURLになることがある');
  } catch (e) { /* 用意できなくても表示は続ける */ }

  var info = webAppUrlInfo_();

  if (!info.url) {
    ui.alert('URLがまだありません',
      'Apps Script エディタの「デプロイ ▸ 新しいデプロイ ▸ ウェブアプリ」で' + LF +
      '　次のユーザーとして実行: 自分' + LF + '　アクセスできるユーザー: 全員' + LF +
      'を選んでデプロイし、できたURLを「' + SH.CONFIG + '」シートの「保護者用URL」に貼り付けてください。',
      ui.ButtonSet.OK);
    return;
  }

  var adminUrl = info.url + (info.url.indexOf('?') >= 0 ? '&' : '?') + 'page=admin';
  syncAdminUrl_();
  ui.showModalDialog(urlDialogHtml_(info.url, adminUrl, info.fromConfig, added), '保護者用URL');
}

/**
 * 「管理画面URL」を設定シートに書き出す。
 * 担任用の入口は保護者用URLに ?page=admin を付けただけだが、
 * それを知らないと開けない。控えとしてシートに残しておく。
 */
function syncAdminUrl_() {
  try {
    // 「保護者用URL」が空のときは何も書かない。
    // 他校へ配るひな形で、こちらのデプロイURLが復活してしまうのを防ぐ
    if (!getConfig().parentUrl) return;

    var info = webAppUrlInfo_();
    if (!info.url) return;
    var adminUrl = info.url + (info.url.indexOf('?') >= 0 ? '&' : '?') + 'page=admin';
    ensureConfigKey_(ADMIN_URL_KEY,
      '担任用の管理画面のURL。保護者用URLから自動で作る。保護者には配らないこと', adminUrl);
    setConfigValue_(ADMIN_URL_KEY, adminUrl);
  } catch (e) {
    console.warn('管理画面URLの書き出しをスキップ:', e);
  }
}

/** URL表示のダイアログ。リンクとコピーボタンを付ける。 */
function urlDialogHtml_(parentUrl, adminUrl, fromConfig, addedRow) {
  var rows = [
    { label: '保護者用（予約ページ）', url: parentUrl },
    { label: '担任用（Web管理画面）', url: adminUrl }
  ];

  var body = '';
  for (var i = 0; i < rows.length; i++) {
    var u = escHtml_(rows[i].url);
    body +=
      '<div class="row">' +
      '<div class="lbl">' + escHtml_(rows[i].label) + '</div>' +
      '<div><a href="' + u + '" target="_blank" rel="noopener">' + u + '</a></div>' +
      '<div class="tools">' +
      '<input id="f' + i + '" readonly value="' + u + '">' +
      '<button onclick="cp(' + i + ',this)">コピー</button>' +
      '</div></div>';
  }

  // 自動取得のときは、貼り付け先までの手順をそのまま出す。
  // 文章で長々と書くと枠に収まらず、肝心の手順が読めなくなる。
  var note = fromConfig
    ? '<b>✓ 設定シートの「保護者用URL」に入力されたURLです。</b>' +
      '案内プリントのQRコードも同じURLで作られます。'
    : '<b>⚠ このURLは自動で取得したものです。実際に配るURLと違うことがあります。</b>' +
      '<ol>' +
      '<li>この画面を閉じて「' + escHtml_(SH.CONFIG) + '」シートを開く</li>' +
      '<li>A列の <b>保護者用URL</b> の行' +
      (addedRow ? '（いま一番下に用意しました）' : '') +
      'を探し、<b>その右となりのセル（B列）</b>に正しいURLを貼る</li>' +
      '<li>もう一度この画面を開く</li>' +
      '</ol>' +
      '正しいURLは Apps Script エディタの「デプロイ ▸ デプロイを管理」で確認できます。' +
      '<b>貼り付けるまで、案内プリントのQRコードも上のURLで作られます。</b>';

  var html =
    '<style>' +
    'body{font-family:"Yu Gothic",Meiryo,sans-serif;font-size:13px;color:#202124;margin:14px}' +
    '.row{margin-bottom:20px}' +
    '.lbl{font-weight:700;margin-bottom:5px}' +
    'a{color:#1a73e8;word-break:break-all;line-height:1.6}' +
    '.tools{display:flex;gap:6px;margin-top:7px}' +
    'input{flex:1;font-size:12px;padding:5px;border:1px solid #dadce0;border-radius:4px;color:#5f6368}' +
    'button{padding:5px 12px;border:1px solid #1a73e8;background:#fff;color:#1a73e8;' +
    'border-radius:4px;cursor:pointer;white-space:nowrap}' +
    '.note{background:' + (fromConfig ? '#e6f4ea' : '#fef7e0') + ';padding:12px 14px;' +
    'border-radius:6px;line-height:1.8;font-size:12px}' +
    '.note ol{margin:8px 0;padding-left:1.3em}' +
    '.note li{margin-bottom:4px}' +
    '</style>' +
    body +
    '<div class="note">' + note + '</div>' +
    '<script>' +
    'function cp(i,b){var f=document.getElementById("f"+i);f.select();' +
    'try{document.execCommand("copy");b.textContent="コピーしました";' +
    'setTimeout(function(){b.textContent="コピー";},1500);}catch(e){}}' +
    '</' + 'script>';

  // 自動取得のときは手順のぶん背が高くなる。切れると手順が読めないので余裕を持たせる。
  return HtmlService.createHtmlOutput(html).setWidth(580).setHeight(fromConfig ? 340 : 500);
}

/* ---------------- Web管理画面 API ---------------- */

function requireAdmin_(pass) {
  var cfg = getConfig();
  if (!cfg.adminPasscode) throw new Error('管理パスコードが未設定です。「設定」シートの「管理パスコード」を入力してください。');
  if (String(pass || '') !== cfg.adminPasscode) {
    Utilities.sleep(800);
    throw new Error('パスコードが違います。');
  }
  return cfg;
}

function apiAdminInit(pass) {
  return safe_(function () {
    var cfg = requireAdmin_(pass);
    var conflicts = [];
    try {
      conflicts = findNgConflicts_(readNgSet_(), readSlots_()).map(function (c) {
        return ngConflictLabel_(c);
      });
    } catch (e) {
      console.warn('だめなコマの確認をスキップ:', e);
    }
    var struggles = [];
    try {
      struggles = recentInputErrors_(14).map(function (x) { return inputErrorLabel_(x); });
    } catch (e) {
      console.warn('入力エラーの集計をスキップ:', e);
    }

    // 予備の枠に入れる生徒をプルダウンで選べるよう、クラスごとの名簿を渡す。
    // この画面はパスコードで守られているので、名簿を出してよい
    var roster = getRoster();
    var rosterByClass = {};
    for (var i = 0; i < roster.length; i++) {
      var k = roster[i].cls;
      if (!rosterByClass[k]) rosterByClass[k] = [];
      rosterByClass[k].push({ no: roster[i].no, name: roster[i].name });
    }

    return {
      title: cfg.title,
      published: cfg.published,
      classes: getClasses(),
      roster: rosterByClass,
      days: getDays().map(function (d) { return { key: ymd_(d), label: dateLabel_(d) }; }),
      conflicts: conflicts,
      struggles: struggles
    };
  });
}

/** 指定日の全クラス分の枠（担任向けなので氏名込み） */
function apiAdminBoard(pass, dayKey) {
  return safe_(function () {
    requireAdmin_(pass);
    var classes = getClasses().map(function (c) { return c.name; });
    var slots = readSlots_();
    var ngSet = {};
    try { ngSet = readNgSet_(); } catch (e) { ngSet = {}; }
    var rows = {};
    var order = [];

    for (var i = 0; i < slots.length; i++) {
      var v = slots[i].v;
      if (dayKey && ymd_(v[COL.DATE - 1]) !== dayKey) continue;

      // 予備は「予約表_〇組」で記入する枠。ここでは中身を見せるだけで、押しても変えられない
      var isReserve = String(v[COL.STATUS - 1]) === STATUS.RESERVE;
      var time = String(v[COL.START - 1]) + '–' + String(v[COL.END - 1]) +
        (isReserve ? '（予備）' : '');
      var key = ymd_(v[COL.DATE - 1]) + ' ' + time;
      if (!rows[key]) {
        rows[key] = { dateLabel: dateLabel_(v[COL.DATE - 1]), time: time, cells: {} };
        order.push(key);
      }
      rows[key].cells[String(v[COL.CLASS - 1])] = {
        slotId: String(v[COL.SLOT_ID - 1]),
        status: String(v[COL.STATUS - 1]),
        reserve: isReserve,
        ng: !!ngSet[String(v[COL.SLOT_ID - 1])],
        no: v[COL.NUMBER - 1],
        name: String(v[COL.STUDENT - 1] || ''),
        guardian: String(v[COL.GUARDIAN - 1] || ''),
        note: String(v[COL.NOTE - 1] || '')
      };
    }
    order.sort();
    return { classes: classes, rows: order.map(function (k) { return rows[k]; }) };
  });
}

function apiAdminSetStatus(pass, slotId, status) {
  return safe_(function () {
    requireAdmin_(pass);
    if (status !== STATUS.OPEN && status !== STATUS.BLOCKED) throw new Error('指定できない状態です。');
    return withLock_(function () {
      var found = findSlotRow_(slotId);
      if (String(found.v[COL.STATUS - 1]) === STATUS.BOOKED) {
        throw new Error('予約が入っています。先に取り消してください。');
      }
      if (String(found.v[COL.STATUS - 1]) === STATUS.RESERVE) {
        throw new Error('予備の枠はここでは変更できません。黄色いセルを押して、生徒を入れるか空にしてください。');
      }
      sheet_(SH.SLOTS).getRange(found.row, COL.STATUS).setValue(status);
      clearSlotCache_();
      markViewsStale_();
      try {
        setNgFlag_(slotId, status === STATUS.BLOCKED);
      } catch (e) {
        console.warn('だめなコマシートの同期をスキップ:', e);
      }
      logAction_(status === STATUS.BLOCKED ? 'ブロック' : 'ブロック解除', slotId, String(found.v[COL.CLASS - 1]), '', '', '担任Web');
      return {};
    });
  });
}

/**
 * 選んだ枠をまとめて「ブロック(×)」または「空き」にする。
 * 「だめなコマ」シートのチェックも同時に更新する。
 */
function apiAdminSetStatusBulk(pass, slotIds, status) {
  return safe_(function () {
    requireAdmin_(pass);
    if (status !== STATUS.OPEN && status !== STATUS.BLOCKED) throw new Error('指定できない状態です。');

    var ids = [];
    var list = slotIds || [];
    for (var i = 0; i < list.length; i++) {
      var id = String(list[i] || '').trim();
      if (id) ids.push(id);
    }
    if (!ids.length) throw new Error('枠が選ばれていません。');

    var res = setStatusForSlotIds_(ids, status);
    logAction_(status === STATUS.BLOCKED ? 'ブロック' : 'ブロック解除', '', '', '', '',
      '担任Web 一括: ' + res.changed + '件' + (res.skipped ? ' / 予約済のため据え置き ' + res.skipped + '件' : ''));
    return {
      changed: res.changed,
      skipped: res.skipped,
      conflicts: (res.skippedList || []).map(function (c) { return ngConflictLabel_(c); })
    };
  });
}

/** まだ予約が入っていない生徒の一覧 */
function apiAdminUnbooked(pass) {
  return safe_(function () {
    requireAdmin_(pass);
    var slots = readSlots_();
    var booked = bookedKeySet_(slots);   // 交流学級での予約も予約済みとして数える

    var roster = getRoster();
    var out = [];
    for (var r = 0; r < roster.length; r++) {
      if (booked[roster[r].cls + '_' + roster[r].no]) continue;
      out.push({ cls: roster[r].cls, no: roster[r].no, name: roster[r].name });
    }
    out.sort(function (a, b) {
      if (a.cls !== b.cls) return a.cls < b.cls ? -1 : 1;
      return a.no - b.no;
    });
    return { students: out };
  });
}

/**
 * 予備の枠に生徒を入れる／外す。
 * 入力は出席番号でも氏名でもよい。名簿と突き合わせて、もう一方を補う。
 * 空文字を渡すと、その枠を空にする。
 */
function apiAdminSetReserve(pass, slotId, input) {
  return safe_(function () {
    requireAdmin_(pass);
    return setReserveStudent_(slotId, input, '担任Web');
  });
}

/**
 * 通常の枠に、担任が代理で生徒を入れる。
 * 電話で申し込まれた場合など、保護者に代わって登録するときに使う。
 *
 * 予約コードを発行して返すので、保護者にはそれを伝える。
 * 伝えておけば、あとの変更・取消は保護者自身の画面からできる。
 */
function apiAdminAssign(pass, slotId, no) {
  return safe_(function () {
    requireAdmin_(pass);
    return assignStudentToSlot_(slotId, no);
  });
}

/** クラス内で、出席番号または氏名から生徒を1人に決める */
function resolveStudent_(cls, text) {
  var t = String(text == null ? '' : text).trim();
  if (!t) throw new Error('出席番号か氏名を入力してください。');

  var roster = getRoster();
  var asNo = Number(t);
  for (var i = 0; i < roster.length; i++) {
    if (roster[i].cls !== cls) continue;
    if (asNo && roster[i].no === asNo) return roster[i];
    if (norm_(roster[i].name) === norm_(t)) return roster[i];
  }
  throw new Error(cls + ' の名簿に「' + t + '」が見つかりません。' +
    '出席番号か、名簿どおりの氏名で入力してください。');
}

/**
 * 通常の枠に、担任が代理で生徒を入れる。
 * Web管理画面とシートのメニューの両方から使う。
 *
 * 電話で申し込まれた場合など、保護者に代わって登録するときのもの。
 * 予約コードを発行して返すので、保護者にはそれを伝える。
 * 伝えておけば、あとの変更・取消は保護者自身の画面からできる。
 */
function assignStudentToSlot_(slotId, text) {
  var pending = null;

  var result = withLock_(function () {
    var found = findSlotRow_(slotId);
    var st = String(found.v[COL.STATUS - 1]);
    if (st === STATUS.BOOKED) {
      throw new Error('この枠にはすでに ' + found.v[COL.NUMBER - 1] + '. ' +
        found.v[COL.STUDENT - 1] + ' さんの予約が入っています。先に取り消してください。');
    }
    if (st === STATUS.RESERVE) {
      throw new Error('予備の枠です。黄色い行の出席番号・生徒氏名・保護者氏名へ直接入力するか、担任用の管理画面を使ってください。');
    }

    var cls = String(found.v[COL.CLASS - 1]);
    var hit = resolveStudent_(cls, text);

    var slots = readSlots_();
    var already = findExistingBookingFor_(slots, cls, hit.no, linkedIdentity_(cls, hit.no));
    if (already) {
      throw new Error(hit.name + ' さんは ' + dateLabel_(already.v[COL.DATE - 1]) + ' ' +
        already.v[COL.START - 1] + ' に面談が入っています。先にそちらを取り消してください。');
    }

    var code = makeCode_();
    sheet_(SH.SLOTS).getRange(found.row, COL.STATUS, 1, SLOT_LAST_COL - COL.STATUS + 1)
      .setValues([[STATUS.BOOKED, hit.no, hit.name, '', '担任が代理で登録', code, new Date()]]);
    clearSlotCache_();
    markViewsStale_();

    // 面談を入れる以上、「だめなコマ」の指定は外しておく
    try {
      setNgFlag_(slotId, false);
    } catch (e) {
      console.warn('だめなコマシートの同期をスキップ:', e);
    }

    logAction_('予約', slotId, cls, hit.no, hit.name, '代理登録 / コード ' + code);
    pending = { slot: found, code: code, student: { cls: cls, no: hit.no, name: hit.name } };

    return {
      no: hit.no, name: hit.name, code: code, cls: cls,
      dateLabel: dateLabel_(found.v[COL.DATE - 1]),
      start: String(found.v[COL.START - 1]),
      end: String(found.v[COL.END - 1])
    };
  });

  if (pending) {
    notifyTeacherAfterLock_('予約', pending.student,
      { slot: pending.slot, guardian: '', note: '担任が代理で登録', code: pending.code });
  }
  return result;
}

/**
 * 予備の枠に生徒を入れる／外す。
 * Web管理画面とシートのメニューの両方から使う。
 *
 * 通常の枠への代理登録と違い、予約コードは発行しない。
 * 予備は保護者の画面に一覧されず、コードで引くこともできない枠だからである。
 *
 * @param {string} slotId 対象の枠ID
 * @param {string} input 出席番号か氏名。空文字ならその枠を空にする
 * @param {string} source 予約ログに残す操作元
 * @param {string=} guardian 保護者氏名。省略時は現在の値を保つ
 * @param {number=} lockWaitMs ロックを待つ時間。単純onEditからは短く指定する
 * @return {{no:(number|string), name:string, cls:string, dateLabel:string, start:string, end:string}}
 */
function setReserveStudent_(slotId, input, source, guardian) {
  var text = String(input == null ? '' : input).trim();
  var from = source || '担任';
  var guardianProvided = arguments.length >= 4;
  var guardianText = guardianProvided ? String(guardian == null ? '' : guardian).trim() : '';
  var lockWaitMs = arguments.length >= 5 ? Number(arguments[4]) : 30000;

  return withLock_(function () {
    var found = findSlotRow_(slotId);
    if (String(found.v[COL.STATUS - 1]) !== STATUS.RESERVE) {
      throw new Error('予備の枠ではありません。');
    }
    var cls = String(found.v[COL.CLASS - 1]);
    var sh = sheet_(SH.SLOTS);
    var when = {
      cls: cls,
      dateLabel: dateLabel_(found.v[COL.DATE - 1]),
      start: String(found.v[COL.START - 1]),
      end: String(found.v[COL.END - 1])
    };

    if (!text) {
      if (!found.v[COL.NUMBER - 1] && !found.v[COL.STUDENT - 1] &&
          !found.v[COL.GUARDIAN - 1] && !found.v[COL.NOTE - 1]) {
        syncReserveRowInClassSheet_(slotId, cls, '', '', '');
        return {
          no: '', name: '', cls: cls,
          dateLabel: when.dateLabel, start: when.start, end: when.end
        };
      }
      sh.getRange(found.row, COL.NUMBER, 1, 4).setValues([['', '', '', '']]);
      clearSlotCache_();
      markViewsStale_();
      syncReserveRowInClassSheet_(slotId, cls, '', '', '');
      logAction_('予備コマを空に', slotId, cls,
        found.v[COL.NUMBER - 1] || '', String(found.v[COL.STUDENT - 1] || ''), from);
      return {
        no: '', name: '', cls: cls,
        dateLabel: when.dateLabel, start: when.start, end: when.end
      };
    }

    var hit = resolveStudent_(cls, text);

    // すでにどこかで面談が入っている生徒を、重ねて予備に入れない
    var slots = readSlots_();
    var already = findExistingBookingFor_(slots, cls, hit.no, linkedIdentity_(cls, hit.no));
    if (already && String(already.v[COL.SLOT_ID - 1]) !== slotId) {
      throw new Error(hit.name + ' さんは ' + dateLabel_(already.v[COL.DATE - 1]) + ' ' +
        already.v[COL.START - 1] + ' に面談が入っています。先にそちらを取り消してください。');
    }

    var savedGuardian = guardianProvided
      ? guardianText
      : String(found.v[COL.GUARDIAN - 1] || '');
    if (Number(found.v[COL.NUMBER - 1]) === Number(hit.no) &&
        String(found.v[COL.STUDENT - 1] || '') === String(hit.name) &&
        String(found.v[COL.GUARDIAN - 1] || '') === savedGuardian) {
      syncReserveRowInClassSheet_(slotId, cls, hit.no, hit.name, savedGuardian);
      return {
        no: hit.no, name: hit.name, cls: cls,
        dateLabel: when.dateLabel, start: when.start, end: when.end
      };
    }

    sh.getRange(found.row, COL.NUMBER, 1, 3)
      .setValues([[hit.no, hit.name, savedGuardian]]);
    clearSlotCache_();
    markViewsStale_();
    syncReserveRowInClassSheet_(slotId, cls, hit.no, hit.name, savedGuardian);
    logAction_('予備コマに記入', slotId, cls, hit.no, hit.name, from);
    return {
      no: hit.no, name: hit.name, cls: cls,
      dateLabel: when.dateLabel, start: when.start, end: when.end
    };
  }, lockWaitMs);
}

/** 枠マスタの予備コマを、該当クラスの表示行へすぐ書き写す。 */
function syncReserveRowInClassSheet_(slotId, cls, no, name, guardian) {
  try {
    var sh = ss_().getSheetByName(CLASS_SHEET_PREFIX + cls);
    if (!sh || sh.getLastRow() < 2) return;
    var idCol = 9 + CLASS_HEADER_RIGHT.length;
    var ids = sh.getRange(2, idCol, sh.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0] || '').trim() !== String(slotId)) continue;
      sh.getRange(i + 2, 12, 1, 3).setValues([[no || '', name || '', guardian || '']]);
      return;
    }
  } catch (e) {
    // 表示への書き写しに失敗しても、枠マスタへの保存は完了している
    console.warn('予備コマの表示同期をスキップ:', e);
  }
}

function apiAdminCancel(pass, slotId) {
  return safe_(function () {
    requireAdmin_(pass);
    return withLock_(function () {
      var found = findSlotRow_(slotId);
      if (String(found.v[COL.STATUS - 1]) !== STATUS.BOOKED) throw new Error('その枠に予約はありません。');
      clearSlotRow_(found.row, slotId);
      clearSlotCache_();
      markViewsStale_();
      logAction_('取消', slotId, String(found.v[COL.CLASS - 1]), found.v[COL.NUMBER - 1],
        String(found.v[COL.STUDENT - 1]),
        '担任Webによる取消 / コード ' + String(found.v[COL.CODE - 1] || ''));
      return {};
    });
  });
}

function apiAdminRefreshViews(pass) {
  return safe_(function () {
    requireAdmin_(pass);
    var pending = pendingViewUpdates_();
    refreshViews(true);
    return { pending: pending };
  });
}

function findSlotRow_(slotId) {
  var slots = readSlots_();
  for (var i = 0; i < slots.length; i++) {
    if (String(slots[i].v[COL.SLOT_ID - 1]) === String(slotId)) return slots[i];
  }
  throw new Error('枠が見つかりません: ' + slotId);
}


/**
 * 「予約表_〇組」の、作り直されてしまう場所に書き込まれたときに知らせる。
 *
 * このシートで手入力してよいのは、A・B列（名簿）と黄色い予備行のL〜N列。
 * それ以外は次の更新で消えるが、書いた本人には分からない。
 * 黙って消すのがいちばん困るので、その場で伝える。
 *
 * 正しい予備入力に警告を出すと、やがて警告の中身を読まずに閉じるようになるため、
 * 許可する列と行を明確に判定する。
 */
function warnGeneratedEdit_(e, sh) {
  var row = e.range.getRow();
  if (row < 2) return;   // 見出し行

  var col = e.range.getColumn();
  var lastCol = col + e.range.getNumColumns() - 1;
  if (lastCol <= 2) return;   // A・B列は名簿。手で入れる場所

  sh.getParent().toast(
    'この表は「' + SH.SLOTS + '」から自動で作り直されます。いま書いた内容は次の更新で消えます。' +
    '手で入れてよいのはA・B列の名簿と、黄色い予備行の出席番号・生徒氏名・保護者氏名だけです。' +
    '面談を入れるときは、その行を選んでメニュー「' + MENU.ROWOPS +
    ' ▸ この枠に生徒を入れる」か、担任用の管理画面から行ってください。',
    '⚠ ここに書いても残りません', 15);

  // どの時間の行かが分かると、あとから担任に確認しやすい
  var where = '';
  try {
    var head = sh.getRange(row, 9, 1, 3).getValues()[0];   // 日付・時間・状態
    if (head[0]) where = String(head[0]) + ' ' + String(head[1]) + '（' + String(head[2]) + '）　';
  } catch (e2) { /* 位置が取れなくても知らせる */ }

  logStrayEdit_(sh.getName().slice(CLASS_SHEET_PREFIX.length), e, where +
    '通常枠の登録は担任用の管理画面か、その行を選んで' +
    '「' + MENU.ROWOPS + ' ▸ この枠に生徒を入れる」から行ってください');
}

/**
 * 黄色い予備行の L〜N列へ直接入力された内容を、その場で枠マスタへ保存する。
 * 名簿照合・二重予約確認は、管理画面と同じ setReserveStudent_ を必ず通す。
 * @return {boolean} 予備行の正規入力として扱ったとき true
 */
function handleReserveEntryEdit_(e, sh) {
  var range = e.range;
  var firstCol = range.getColumn();
  var lastCol = firstCol + range.getNumColumns() - 1;
  if (firstCol < 12 || lastCol > 14 || range.getRow() < 2) return false;

  var rowCount = range.getNumRows();
  var rows = sh.getRange(range.getRow(), 11, rowCount, 6).getValues(); // K〜P
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) !== STATUS.RESERVE || !String(rows[i][5] || '').trim()) return false;
  }

  var touchesNo = firstCol <= 12 && lastCol >= 12;
  var touchesName = firstCol <= 13 && lastCol >= 13;
  for (var r = 0; r < rows.length; r++) {
    var sheetRow = range.getRow() + r;
    var slotId = String(rows[r][5]).trim();
    var typedNo = String(rows[r][1] == null ? '' : rows[r][1]).trim();
    var typedName = String(rows[r][2] || '').trim();
    var typedGuardian = String(rows[r][3] || '').trim();
    try {
      var input = '';
      if (touchesNo && touchesName && typedNo && typedName) {
        var byNo = resolveStudent_(sh.getName().slice(CLASS_SHEET_PREFIX.length), typedNo);
        var byName = resolveStudent_(sh.getName().slice(CLASS_SHEET_PREFIX.length), typedName);
        if (Number(byNo.no) !== Number(byName.no)) {
          throw new Error('出席番号と生徒氏名が一致しません。');
        }
        input = String(byNo.no);
      } else if (touchesName && typedName) {
        input = typedName;
      } else if (touchesNo && typedNo) {
        input = typedNo;
      } else {
        input = typedNo || typedName;
      }

      if (!input && typedGuardian) {
        throw new Error('先に出席番号か生徒氏名を入力してください。');
      }
      // 単純 onEdit は最長30秒なので、保護者予約と重なったときは早めに戻して
      // 担任へ再入力を案内する。Web・メニュー操作は従来どおり30秒待つ。
      setReserveStudent_(slotId, input, '担任が予約表に直接入力', typedGuardian, 3000);
    } catch (err) {
      try {
        var found = findSlotRow_(slotId);
        sh.getRange(sheetRow, 12, 1, 3).setValues([[
          found.v[COL.NUMBER - 1] || '',
          found.v[COL.STUDENT - 1] || '',
          found.v[COL.GUARDIAN - 1] || ''
        ]]);
      } catch (restoreErr) { /* 次の表示更新でも元に戻る */ }
      sh.getParent().toast(String(err.message || err), '⚠ 予備枠に登録できません', 15);
      logStrayEdit_(sh.getName().slice(CLASS_SHEET_PREFIX.length), e,
        '予備枠への入力を確認してください: ' + String(err.message || err));
    }
  }
  return true;
}

/**
 * 「全体ビュー」に手で書き込まれたときに知らせる。
 * こちらは丸ごと作り直すので、残る場所が1つも無い。
 */
function warnOverviewEdit_(e, sh) {
  sh.getParent().toast(
    'この表は「' + SH.SLOTS + '」から自動で作り直されます。いま書いた内容は次の更新で消えます。' +
    '面談を入れない時間の指定は「' + SH.NG + '」シートか、担任用の管理画面で行ってください。',
    '⚠ ここに書いても残りません', 15);

  logStrayEdit_('', e,
    '面談を入れない時間の指定は「' + SH.NG + '」シートか、担任用の管理画面で行ってください');
}

/**
 * 登録されないまま消える手入力を、予約ログへ残す。
 *
 * トーストは15秒で流れてしまい、見落とすと誰にも伝わらない。
 * 「入れたつもり」のまま当日を迎えるのがいちばん困るので、必ず記録に残し、
 * 🩺データ点検と受付開始前の点検から拾えるようにする。
 */
function logStrayEdit_(cls, e, advice) {
  try {
    var at = '';
    try { at = e.range.getSheet().getName() + ' の ' + e.range.getA1Notation(); } catch (e2) { at = ''; }
    logAction_(STRAY_EDIT_ACTION, '', cls || '', '', '',
      (at ? at + ' ' : '') + 'に書き込まれましたが、この場所は自動で作り直されるため登録されていません' +
      (e && e.value != null ? '（入力: ' + String(e.value) + '）' : '') +
      (advice ? ' ／ ' + advice : ''));
  } catch (err) {
    // 記録できなくても、トーストは出ている。編集そのものは邪魔しない
    console.warn('未登録の手入力を記録できませんでした:', err);
  }
}

/* ---------------- 受付状態の表示 ---------------- */

/**
 * 「設定」シートで受付の可否を手で書き換えたときに、メニューの表示を追従させる。
 *
 * メニューはスプレッドシートを開いた瞬間にしか組み立てられないので、
 * 「公開」を手で TRUE にしてもラベルは「🔴 停止中」のまま残ってしまっていた。
 * 単純トリガーなので、関係のない編集では何もせずに抜ける。
 */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    var sh = e.range.getSheet();
    var sheetName = sh.getName();

    // 参照データを書き換えたら短時間キャッシュを捨て、次の読み取りから反映させる
    if (sheetName === SH.CONFIG || sheetName === SH.CLASSES || sheetName === SH.LINK ||
        sheetName.indexOf(CLASS_SHEET_PREFIX) === 0) {
      dropRefCaches_();
    }

    // 予約表・全体ビューは毎回作り直される表示用のシート。
    // 書いても残らない場所に手を入れたら、その場で知らせる
    if (sheetName.indexOf(CLASS_SHEET_PREFIX) === 0) {
      if (handleReserveEntryEdit_(e, sh)) return;
      warnGeneratedEdit_(e, sh);
      return;
    }
    if (sheetName === SH.OVERVIEW) {
      warnOverviewEdit_(e, sh);
      return;
    }

    if (sheetName !== SH.CONFIG) return;

    var keys = sh.getRange(e.range.getRow(), 1, e.range.getNumRows(), 1).getValues();
    for (var i = 0; i < keys.length; i++) {
      var k = String(keys[i][0] || '').trim();

      // 自動処理の行のトリガー付け外しは onConfigEdit（取付トリガー）が行う。
      // ここは単純トリガーなので、メニューの表示を追従させるだけ。
      if (k === '公開' || k === '予約受付開始' || k === '予約受付締切' ||
          k === NG_AUTO_KEY || k === VIEW_AUTO_KEY ||
          k === REMINDER_AUTO_KEY || k === BACKUP_AUTO_KEY) {
        buildMenu_();
        return;
      }
    }
  } catch (err) {
    // onEdit はメニュー更新だけでなく、黄色い予備行の保存も担当している。
    // 失敗を黙って飲み込むと「入力したつもり」になるため、担任へ必ず知らせる。
    try {
      var failedSheet = e && e.range ? e.range.getSheet() : null;
      var failedName = failedSheet ? failedSheet.getName() : '';
      if (failedSheet && failedName.indexOf(CLASS_SHEET_PREFIX) === 0) {
        failedSheet.getParent().toast(
          '予備枠への入力を保存できませんでした。表示を更新してから、もう一度入力してください。',
          '⚠ 入力を保存できません', 15);
        logStrayEdit_(failedName.slice(CLASS_SHEET_PREFIX.length), e,
          '予備枠の保存処理でエラー: ' + String(err.message || err));
      }
    } catch (notifyErr) { /* 通知にも失敗した場合だけコンソールへ残す */ }
    console.warn('onEdit の処理に失敗:', err);
  }
}


/**
 * 受付がオンかオフかは、メニューを開いただけでは分からなかった。
 * サブメニューの名前そのものに、いまの状態を出す。
 * 「設定」シートがまだ無い段階でもメニュー自体は開けるようにしておく。
 */
/**
 * 自動処理のメニュー項目名に、いまオンかオフかを付ける。
 * 状態は「設定」シートに書き写した値を読む
 * （onOpen は認可が無く、トリガーそのものを見に行けないため）。
 */
function autoLabel_(flags, key, label) {
  return label + (flags[key] ? '（🟢 オン）' : '（⚪ オフ）');
}

function ngMenuLabel_(flags) {
  return MENU.NG + (flags[NG_AUTO_KEY] ? '（🟢 自動反映オン）' : '（⚪ 手動）');
}

function receptionMenuLabel_() {
  try {
    var cfg = getConfig();
    if (!cfg.published) return MENU.RECEPTION + '（🔴 停止中）';
    var win = bookingWindow_(cfg);
    if (!win.ok) return MENU.RECEPTION + '（🟡 期間外）';
    return MENU.RECEPTION + '（🟢 受付中）';
  } catch (e) {
    return MENU.RECEPTION;
  }
}

/** いま保護者が予約できる状態かどうかを、理由つきで表示する。 */
function menuReceptionStatus() {
  var ui = SpreadsheetApp.getUi();
  try {
    var cfg = getConfig();
    var win = bookingWindow_(cfg);
    var lines = [];

    lines.push(win.ok ? '🟢 いま保護者は予約できます。' : '🔴 いま保護者は予約できません。');
    if (!win.ok) lines.push('　理由: ' + win.message);

    lines.push('');
    lines.push('「' + SH.CONFIG + '」シートの値');
    lines.push('　公開: ' + (cfg.published ? 'TRUE（受付する）' : 'FALSE（受付しない）'));
    lines.push('　予約受付開始: ' +
      (cfg.openAt ? Utilities.formatDate(cfg.openAt, TZ, 'yyyy/MM/dd HH:mm') : '（空欄＝すぐ受付）'));
    lines.push('　予約受付締切: ' +
      (cfg.closeAt ? Utilities.formatDate(cfg.closeAt, TZ, 'yyyy/MM/dd HH:mm') : '（空欄＝締切なし）'));

    try {
      var slots = readSlots_();
      var booked = 0;
      for (var i = 0; i < slots.length; i++) {
        if (isTakenSlot_(slots[i].v)) booked++;
      }
      lines.push('');
      lines.push('予約ずみ: ' + getRoster().length + '名中 ' + booked + '名');
    } catch (e) { /* 件数が出せなくても受付状態は伝わる */ }

    lines.push('');
    lines.push('システムのバージョン: ' + VERSION);

    ui.alert('予約の受付状態', lines.join(LF), ui.ButtonSet.OK);
  } catch (err) {
    ui.alert('受付状態を確認できません', String(err.message || err), ui.ButtonSet.OK);
  }
}
