/**
 * 自動化まわり。
 *
 *  1. だめなコマのチェックを、入力した瞬間に枠へ反映する（編集トリガー）
 *  2. 締切前に、未予約の生徒一覧を担任へメールで知らせる（日次トリガー）
 *
 * どちらも既定では止まっている。メニューから明示的に有効化してもらう。
 */

var NG_TRIGGER_FN = 'onNgEdit';
var REMINDER_TRIGGER_FN = 'dailyReminderCheck';
var BACKUP_TRIGGER_FN = 'dailyBackupIfOpen';
var CONFIG_TRIGGER_FN = 'onConfigEdit';

/* ================================================================
   1. だめなコマの自動反映
   ================================================================ */

/**
 * トリガーの有無（＝実態）を「設定」シートへ書き写す。
 *
 * メニューを組み立てる onOpen は単純トリガーで、認可の要る
 * ScriptApp.getProjectTriggers() を呼べない。そのため実態はメニュー生成時には読めない。
 * 認可のある実行から呼ばれるたびに書き写しておき、ラベルはそちらを読む。
 * 呼ばれるたびに実態へ合わせ直すので、エディタで直接トリガーを消してもいずれ直る。
 */
function syncAutoFlag_(key, on) {
  try {
    ensureConfigKey_(key, AUTO_FLAG_DESC[key] || '', on);
    // 説明の文面を直したときに、すでにある行にも反映させる
    setConfigDescription_(key, AUTO_FLAG_DESC[key]);
    setConfigValue_(key, on);
  } catch (e) {
    console.warn('自動処理の状態を書けませんでした: ' + key, e);
  }
}

/**
 * 「設定」シートの見張り役を取り付ける。
 *
 * 単純トリガーの onEdit は認可が無いのでトリガーを付け外しできない。
 * 取付トリガーなら認可があるため、シートの値をそのまま反映できる。
 * これがある限り、自動処理の4行は「公開」と同じく入力として扱える。
 *
 * @return {boolean} 新しく取り付けたら true
 */
function ensureConfigWatcher() {
  if (countTriggers_(CONFIG_TRIGGER_FN) > 0) return false;
  ScriptApp.newTrigger(CONFIG_TRIGGER_FN).forSpreadsheet(ssId_()).onEdit().create();
  return true;
}

/** 自動処理のキー → 対応するトリガー名と、オン・オフの手続き */
function autoFlagActions_() {
  var m = {};
  m[NG_AUTO_KEY] = {
    fn: NG_TRIGGER_FN,
    on: function () { enableNgAutoApply(); },
    off: function () { disableNgAutoApply(); }
  };
  m[VIEW_AUTO_KEY] = {
    fn: VIEW_TRIGGER_FN,
    on: function () { enableAutoRefresh(5); },
    off: function () { disableAutoRefresh(); }
  };
  m[REMINDER_AUTO_KEY] = {
    fn: REMINDER_TRIGGER_FN,
    on: function () { enableReminder(); },
    off: function () { disableReminder(); }
  };
  m[BACKUP_AUTO_KEY] = {
    fn: BACKUP_TRIGGER_FN,
    on: function () { enableAutoBackup(); },
    off: function () { disableAutoBackup(); }
  };
  return m;
}

/**
 * 「設定」シートで自動処理の行が書き換えられたら、実際にトリガーを付け外しする。
 *
 * 値がすでに実態と同じなら何もしない。オン・オフの手続きの中で
 * 設定シートへ書き戻す処理が走るが、値が変わらないときは書き込まないので、
 * この編集トリガーが自分自身を呼び続けることはない。
 */
function onConfigEdit(e) {
  try {
    if (!e || !e.range) return;
    var sh = e.range.getSheet();

    // 枠マスタを直接直したときも、表を作り直す対象にする。
    // ここは元データなので手で直しても有効だが、印を付けないと表に出てこない
    if (sh.getName() === SH.SLOTS) {
      markViewsStale_();
      clearSlotCache_();
      return;
    }

    if (sh.getName() !== SH.CONFIG) return;

    var actions = autoFlagActions_();
    var rows = sh.getRange(e.range.getRow(), 1, e.range.getNumRows(), 2).getValues();
    var changed = [];

    for (var i = 0; i < rows.length; i++) {
      var key = String(rows[i][0] || '').trim();

      // 保護者用URLを貼り替えたら、そこから作る管理画面URLもその場で作り直す。
      // 古いほうが残っていると、担任がどちらを開けばよいのか分からなくなる
      if (key === '保護者用URL') {
        try {
          syncAdminUrl_();
          changed.push('管理画面URLを作り直しました');
        } catch (e3) {
          console.warn('管理画面URLの更新をスキップ:', e3);
        }
        continue;
      }

      var act = actions[key];
      if (!act) continue;

      var want = truthy_(rows[i][1]);
      if (want === (countTriggers_(act.fn) > 0)) continue;

      if (want) act.on(); else act.off();
      changed.push(key + ' を' + (want ? 'オン' : 'オフ') + 'にしました');
    }

    if (!changed.length) return;
    sh.getParent().toast(changed.join(' / '), '三者面談', 8);

    // 取付トリガーからは UI を触れないことがあるので、失敗しても無視する
    try { buildMenu_(); } catch (e2) { /* メニューは次に開いたときに直る */ }
  } catch (err) {
    console.error('設定シートの変更を反映できませんでした:', err);
  }
}

/**
 * 4つの自動処理の状態を、トリガーを1回だけ読んで「設定」シートへ書き写す。
 * 表示用の行を手で書き換えられたときに、実態へ戻すためにも使う。
 * @return {boolean} 書き写せたら true（認可の無い実行では false）
 */
function syncAllAutoFlags_() {
  var running = {};
  try {
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
      running[triggers[i].getHandlerFunction()] = true;
    }
  } catch (e) {
    return false;
  }
  syncAutoFlag_(NG_AUTO_KEY, !!running[NG_TRIGGER_FN]);
  syncAutoFlag_(VIEW_AUTO_KEY, !!running[VIEW_TRIGGER_FN]);
  syncAutoFlag_(REMINDER_AUTO_KEY, !!running[REMINDER_TRIGGER_FN]);
  syncAutoFlag_(BACKUP_AUTO_KEY, !!running[BACKUP_TRIGGER_FN]);
  return true;
}

/** 自動反映が有効かどうか（トリガーの実態を見る） */
function ngAutoApplyEnabled_() {
  var on = countTriggers_(NG_TRIGGER_FN) > 0;
  syncAutoFlag_(NG_AUTO_KEY, on);
  return on;
}

/** @return {boolean} 新しく有効にしたら true（すでに有効なら false） */
function enableNgAutoApply() {
  if (ngAutoApplyEnabled_()) return false;
  ScriptApp.newTrigger(NG_TRIGGER_FN).forSpreadsheet(ssId_()).onEdit().create();
  syncAutoFlag_(NG_AUTO_KEY, true);
  return true;
}

/** @return {number} 削除したトリガー数 */
function disableNgAutoApply() {
  var n = deleteTriggers_(NG_TRIGGER_FN);
  syncAutoFlag_(NG_AUTO_KEY, false);
  return n;
}

/**
 * 「だめなコマ」シートが編集されたときに走る。
 * チェックした瞬間に保護者の画面へ反映されるので、「反映する」の実行忘れが起きない。
 */
function onNgEdit(e) {
  try {
    if (!e || !e.range) return;

    var sh = e.range.getSheet();
    if (sh.getName() !== SH.NG) return;

    // 見出しや日付列だけの編集では動かさない
    var lastRow = e.range.getRow() + e.range.getNumRows() - 1;
    var lastCol = e.range.getColumn() + e.range.getNumColumns() - 1;
    if (lastRow < NG_FIRST_ROW) return;
    if (lastCol < NG_COL.FIRST_CLASS) return;

    var res = applyNgSlots();
    rebuildOverview();
    rebuildClassSheets();

    var parts = [];
    if (res.blocked) parts.push(res.blocked + '枠を面談なしにしました');
    if (res.unblocked) parts.push(res.unblocked + '枠を空きに戻しました');
    if (res.conflicts.length) {
      parts.push('⚠ 予約が入っている ' + res.conflicts.length + '枠はそのままです');
    }
    if (parts.length) {
      ss_().toast(parts.join('／'), '🚫 だめなコマ 自動反映', 6);
    }
  } catch (err) {
    console.warn('だめなコマの自動反映に失敗:', err);
  }
}

/* ================================================================
   2. 締切前の未予約リマインド
   ================================================================ */

/**
 * クラスごとの未予約者。
 * @return {Array<{cls:string, teacher:string, email:string, students:Array<Object>, total:number}>}
 */
function unbookedByClass_() {
  var classes = getClasses();
  var slots = readSlots_();
  var roster = getRoster();

  // 交流学級と、担任が予備の枠に入れたぶんも「予約あり」として数える
  var booked = bookedKeySet_(slots);

  var out = [];
  for (var c = 0; c < classes.length; c++) {
    var cls = classes[c];
    var mine = [];
    var total = 0;
    for (var r = 0; r < roster.length; r++) {
      if (roster[r].cls !== cls.name) continue;
      total++;
      if (!booked[roster[r].cls + '_' + roster[r].no]) {
        mine.push({ no: roster[r].no, name: roster[r].name });
      }
    }
    mine.sort(function (a, b) { return a.no - b.no; });
    out.push({
      cls: cls.name, teacher: cls.teacher, email: cls.email,
      students: mine, total: total
    });
  }
  return out;
}

/**
 * 未予約の生徒一覧を、各クラスの担任へメールで送る。
 * 未予約が0名のクラスには送らない。
 * @return {{sent:Array<Object>, noEmail:Array<string>, done:Array<string>}}
 */
function sendUnbookedReminder() {
  var cfg = getConfig();
  var groups = unbookedByClass_();

  var sent = [], noEmail = [], done = [];
  var closeText = cfg.closeAt
    ? Utilities.formatDate(cfg.closeAt, TZ, 'M月d日 HH:mm') + ' 締切'
    : '';

  for (var i = 0; i < groups.length; i++) {
    var g = groups[i];
    if (!g.students.length) { done.push(g.cls); continue; }
    if (!g.email) { noEmail.push(g.cls); continue; }

    var lines = g.students.map(function (s) { return '　' + s.no + '. ' + s.name; });
    var subject = '【三者面談】' + g.cls + ' 未予約 ' + g.students.length + '名のお知らせ';
    var body = (g.teacher ? g.teacher + ' 先生\n\n' : '') +
      '三者面談の予約状況をお知らせします。' + (closeText ? '（' + closeText + '）' : '') + '\n\n' +
      '----------------------------------------\n' +
      '■ ' + g.cls + '　' + (g.total - g.students.length) + ' / ' + g.total + ' 名 予約済み\n' +
      '■ 未予約 ' + g.students.length + ' 名\n' +
      lines.join('\n') + '\n' +
      '----------------------------------------\n\n' +
      'お手数ですが、ご家庭への声かけをお願いいたします。\n' +
      '※このメールは「' + SH.CONFIG + '」シートの設定にもとづいて送信しています。';

    try {
      MailApp.sendEmail(g.email, subject, body);
      sent.push({ cls: g.cls, email: g.email, count: g.students.length });
    } catch (err) {
      console.warn('リマインド送信に失敗: ' + g.cls, err);
      noEmail.push(g.cls + '（送信エラー）');
    }
  }

  logAction_('未予約リマインド', '', '', '', '',
    '送信 ' + sent.length + 'クラス / 全員予約済み ' + done.length + 'クラス');

  return { sent: sent, noEmail: noEmail, done: done };
}

/** 自動リマインドが有効かどうか */
function reminderEnabled_() {
  var on = countTriggers_(REMINDER_TRIGGER_FN) > 0;
  syncAutoFlag_(REMINDER_AUTO_KEY, on);
  return on;
}

/** @return {boolean} 新しく有効にしたら true */
function enableReminder() {
  if (reminderEnabled_()) return false;
  ScriptApp.newTrigger(REMINDER_TRIGGER_FN).timeBased().atHour(7).everyDays(1).create();
  syncAutoFlag_(REMINDER_AUTO_KEY, true);
  return true;
}

/** @return {number} 削除したトリガー数 */
function disableReminder() {
  var n = deleteTriggers_(REMINDER_TRIGGER_FN);
  syncAutoFlag_(REMINDER_AUTO_KEY, false);
  return n;
}

/**
 * 毎日1回走り、締切の指定日数前になったらリマインドを送る。
 * 締切が未設定なら何もしない。
 */
function dailyReminderCheck() {
  try {
    var cfg = getConfig();
    if (!cfg.published || !cfg.closeAt) return;

    var target = new Date(cfg.closeAt.getTime());
    target.setDate(target.getDate() - cfg.reminderDays);
    if (ymd_(target) !== ymd_(new Date())) return;

    sendUnbookedReminder();
  } catch (err) {
    console.error('リマインドの自動送信に失敗:', err);
  }
}

/* ================================================================
   トリガーの共通処理
   ================================================================ */

/* ================================================================
   3. 受付期間中の自動バックアップ
   ================================================================ */

function autoBackupEnabled_() {
  var on = countTriggers_(BACKUP_TRIGGER_FN) > 0;
  syncAutoFlag_(BACKUP_AUTO_KEY, on);
  return on;
}

/** @return {boolean} 新しく有効にしたら true */
function enableAutoBackup() {
  if (autoBackupEnabled_()) return false;
  ScriptApp.newTrigger(BACKUP_TRIGGER_FN).timeBased().atHour(2).everyDays(1).create();
  syncAutoFlag_(BACKUP_AUTO_KEY, true);
  return true;
}

/** @return {number} 削除したトリガー数 */
function disableAutoBackup() {
  var n = deleteTriggers_(BACKUP_TRIGGER_FN);
  syncAutoFlag_(BACKUP_AUTO_KEY, false);
  return n;
}

/**
 * 毎日1回走る。受付が開いている間だけバックアップを取る。
 *
 * 予約が消えると復旧できないため、いちばん失いたくない時期を自動で守る。
 * 締め切ったあとや停止中は何もしないので、放っておいてもファイルは増え続けない。
 */
function dailyBackupIfOpen() {
  try {
    var cfg = getConfig();
    if (!cfg.published) return;
    if (cfg.closeAt && new Date() > cfg.closeAt) return;

    var res = createAutoBackup();

    // 受付期間が長いとファイルが溜まり続けるので、古い世代はゴミ箱へ移す。
    // 失敗してもバックアップ自体は取れているので、記録だけ残して続行する。
    var cleaned = 0;
    try {
      cleaned = pruneAutoBackups_(cfg.autoBackupKeep);
    } catch (e) {
      console.warn('古い自動バックアップの整理に失敗:', e);
    }

    logAction_('自動バックアップ', '', '', '', '',
      (res && res.fileName ? res.fileName : '') +
      (cleaned ? ' / 古い ' + cleaned + '件をゴミ箱へ' : ''));
  } catch (err) {
    console.error('自動バックアップに失敗:', err);
  }
}

/* ---------------- トリガーの共通処理 ---------------- */

function countTriggers_(fnName) {
  var n = 0;
  try {
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === fnName) n++;
    }
  } catch (e) {
    console.warn('トリガーの確認に失敗:', e);
  }
  return n;
}

function deleteTriggers_(fnName) {
  var n = 0;
  try {
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === fnName) {
        ScriptApp.deleteTrigger(triggers[i]);
        n++;
      }
    }
  } catch (e) {
    console.warn('トリガーの削除に失敗:', e);
  }
  return n;
}
