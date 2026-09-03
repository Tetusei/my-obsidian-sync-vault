/**
 * リセット（次回にそなえた片づけ）。
 *
 * 消す範囲を間違えると、どちらの方向にも取り返しがつかない。
 *   ・消しすぎる → 名簿が飛ぶ。予約表のA・B列にしか控えが無い
 *   ・消し足りない → 他校へ渡したひな形にこちらの保護者用URLが残り、
 *     コピー先の案内プリントのQRが**この学校の予約ページ**を指す
 *
 * 3段階に分けてあるので、それぞれ「何が消えて、何が残るか」を固定する。
 */

'use strict';

const m = require('./gasmock');

const g = m.load();
m.seedSchool(g, { perClass: 3 });

const C1 = '1年1組';
const C2 = '1年2組';
const DAY1 = '2026-10-29';
const DAY2 = '2026-10-30';

g.setConfigValue_(g.RESERVE_COUNT_KEY, 1);
g.dropRefCaches_();
g.generateSlots();

const A = { cls: C1, no: 1, name: '生徒1_1' };
const B = { cls: C1, no: 2, name: '生徒1_2' };
const C = { cls: C2, no: 1, name: '生徒2_1' };

const cfg = (k) => {
  const sh = g.__ss.getSheetByName(g.SH.CONFIG);
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  const hit = rows.find((r) => String(r[0]).trim() === k);
  return hit ? hit[1] : undefined;
};

function triggerCount(fn) {
  return g.__state.triggers.filter((t) => t.getHandlerFunction() === fn).length;
}

/** 予約・だめなコマ・予備の記入を、ひととおり入れた状態にする */
function fillUp() {
  // リセットは受付を停止状態にするので、入れ直す前に開け直す
  g.setConfigValue_('公開', true);
  g.setConfigValue_('予約受付開始', '');
  g.setConfigValue_('予約受付締切', '');
  g.dropRefCaches_();

  const b1 = m.unwrap(g.apiBook(Object.assign({}, A, { slotId: m.slotIdOf(g, DAY1, C1, 1) })), 'A');
  const b2 = m.unwrap(g.apiBook(Object.assign({}, B, { slotId: m.slotIdOf(g, DAY1, C1, 2) })), 'B');
  const b3 = m.unwrap(g.apiBook(Object.assign({}, C, { slotId: m.slotIdOf(g, DAY1, C2, 1) })), 'C');
  g.setNgFlag_(m.slotIdOf(g, DAY2, C1, 3), true);   // 予約の無い枠を面談なしに
  g.applyNgSlots();
  g.setReserveStudent_(m.slotIdOf(g, DAY1, C1, g.RESERVE_INDEX_BASE + 1), '3', 'テスト');
  return { b1, b2, b3 };
}

/* ---------------- 数える ---------------- */

fillUp();

m.eq(g.countBookings_(), 4, '予約3件と、予備に入れた1件を数える');
m.eq(g.countNgMarks_(), 1, 'だめなコマの指定を数える');
m.eq(g.countRoster_(), 6, '名簿の人数を数える');
m.ok(g.countLog_() > 0, '予約ログの行数を数える');

/* ================================================================
   予約をすべて取り消す
   ================================================================ */

const cleared = g.clearAllBookings_();
m.eq(cleared, 4, '予約3件＋予備1件を取り消したと返す');

m.eq(m.statusOf(g, m.slotIdOf(g, DAY1, C1, 1)), '空き', '予約は空きに戻る');
m.eq(m.statusOf(g, m.slotIdOf(g, DAY1, C2, 1)), '空き', '他クラスの予約も戻る');
m.eq(m.slotValue(g, m.slotIdOf(g, DAY1, C1, 1), g.COL.STUDENT), '', '生徒氏名が消える');
m.eq(m.slotValue(g, m.slotIdOf(g, DAY1, C1, 1), g.COL.CODE), '', '予約コードも消える');

m.eq(m.statusOf(g, m.slotIdOf(g, DAY2, C1, 3)), 'ブロック',
  '**だめなコマに指定した枠は、空きではなくブロックのまま**（面談を入れない設定は残す）');

const rid = m.slotIdOf(g, DAY1, C1, g.RESERVE_INDEX_BASE + 1);
m.eq(m.statusOf(g, rid), '予備', '予備の枠そのものは残る');
m.eq(m.slotValue(g, rid, g.COL.STUDENT), '', '予備に入れた生徒だけが消える');

m.eq(g.countRoster_(), 6, '**名簿には触らない**');
m.eq(g.countNgMarks_(), 1, 'だめなコマの指定にも触らない');

m.eq(g.clearAllBookings_(), 0, '取り消すものが無ければ0件（空の予備枠を数えない）');

const logRows = () => {
  const sh = g.__ss.getSheetByName(g.SH.LOG);
  return sh.getLastRow() < 2 ? [] : sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues();
};
m.ok(logRows().some((r) => String(r[1]) === '一括取消'), '一括取消がログに残る');

/* ================================================================
   だめなコマの指定をすべて外す
   ================================================================ */

m.eq(g.clearAllNg_(), 1, '外した数を返す');
m.eq(g.countNgMarks_(), 0, '指定が無くなる');
m.eq(g.applyNgSlots().unblocked, 1, '反映するとブロックが空きに戻る');
m.eq(m.statusOf(g, m.slotIdOf(g, DAY2, C1, 3)), '空き', '空きに戻っている');

/* ================================================================
   予約ログを消す
   ================================================================ */

const before = g.countLog_();
m.ok(before > 0, '前提: ログがある');
m.eq(g.clearLog_(), before, '消した行数を返す');
m.eq(g.countLog_(), 0, 'ログが空になる');

const logSh = g.__ss.getSheetByName(g.SH.LOG);
m.eq(String(logSh.getRange(1, 1).getValue()), '日時', '見出しは残す');
m.eq(String(logSh.getRange(1, 2).getValue()), '操作', '見出しの並びも残る');

/* ================================================================
   次の面談にそなえて初期化
   ================================================================ */

fillUp();
g.setConfigValue_('予約受付開始', new Date(2026, 9, 1, 9, 0));
g.setConfigValue_('予約受付締切', new Date(2026, 9, 20, 17, 0));
g.setConfigValue_('管理パスコード', 'himitsu');
g.setConfigValue_('保護者用URL', 'https://script.google.com/macros/s/AAA/exec');
g.dropRefCaches_();

const next = g.resetForNext_();

m.eq(next.bookings, 4, '取り消した予約の数を返す');
m.eq(next.ng, 1, '外しただめなコマの数を返す');
m.ok(next.log > 0, '消したログの行数を返す');

/* ---------------- 消えるもの ---------------- */

m.eq(g.countBookings_(), 0, '予約はすべて消える');
m.eq(g.countNgMarks_(), 0, 'だめなコマの指定も消える');
m.eq(cfg('公開'), false, '受付は停止状態になる（うっかり公開したままにしない）');
m.eq(cfg('予約受付開始'), '', '受付開始日時は空になる');
m.eq(cfg('予約受付締切'), '', '締切日時も空になる');
m.eq(m.statusOf(g, m.slotIdOf(g, DAY2, C1, 3)), '空き', 'ブロックも空きに戻る');

/* ---------------- 残るもの ---------------- */

m.eq(g.countRoster_(), 6, '**名簿は残る**（次の面談でも同じ生徒なので）');
m.eq(g.getClasses().map((c) => c.name), [C1, C2], 'クラス名は残る');
m.eq(g.getClasses()[0].teacher, '山田', '担任名も残る');
m.eq(g.getClasses()[0].email, 'y@example.jp', '担任メールも残る');
m.eq(g.getDays().length, 2, '面談日は残る（次回の日付に書き換えて使う）');
m.eq(g.getConfig().slotsPerDay, 3, '1日の枠数は残る');
m.eq(g.getConfig().startMin, 13 * 60 + 40, '面談開始時刻も残る');
m.eq(cfg('管理パスコード'), 'himitsu', '管理パスコードは残る');
m.eq(cfg('保護者用URL'), 'https://script.google.com/macros/s/AAA/exec',
  '保護者用URLも残る（同じ学校で使い続けるため）');
m.ok(g.readSlots_().length > 0, '面談枠そのものは残る');
m.eq(m.statusOf(g, rid), '予備', '予備の枠も残る');

m.ok(logRows().some((r) => String(r[1]) === 'リセット'), 'リセットしたことはログに残る');

/* ================================================================
   年度末：すべて初期化
   ================================================================ */

fillUp();
const all = g.resetAll_();

m.eq(all.roster, 6, '消した名簿の人数を返す');
m.ok(all.slots > 0, '消した枠の数を返す');

/* ---------------- 消えるもの ---------------- */

m.eq(g.countRoster_(), 0, '**名簿が消える**（次の学年は別の生徒なので）');
m.eq(g.getClasses()[0].teacher, '', '担任名が消える');
m.eq(g.getClasses()[0].email, '', '担任メールも消える');
m.eq(g.getDays().length, 0, '面談日が消える');
m.eq(g.readSlots_().length, 0, '面談枠が消える');
m.eq(g.countBookings_(), 0, '予約も消える');
m.eq(g.countLog_(), 0, 'ログも消える');

const ngSh = g.__ss.getSheetByName(g.SH.NG);
m.eq(String(ngSh.getRange(g.NG_FIRST_ROW, 1).getValue()), '', 'だめなコマシートの中身が消える');
m.eq(String(ngSh.getRange(g.NG_BANNER_ROW, 1).getValue()), g.NG_BANNER_TEXT,
  'だめなコマシートの案内は残す');

/* ---------------- 残るもの ---------------- */

m.eq(g.getClasses().map((c) => c.name), [C1, C2],
  '**クラス名は残る**（毎年入れ直す必要が無い）');
m.eq(g.getClasses()[0].grade, '1年', '学年も残る');
m.eq(g.getConfig().slotsPerDay, 3, '面談の時間設定は残る');
m.eq(g.getConfig().slotMin, 15, '枠の長さも残る');
m.eq(cfg('管理パスコード'), 'himitsu', '管理パスコードは残る（同じ学校で使い続けるため）');
m.eq(cfg('保護者用URL'), 'https://script.google.com/macros/s/AAA/exec', '保護者用URLも残る');

const clsShName = g.CLASS_SHEET_PREFIX + C1;
m.ok(g.__ss.getSheetByName(clsShName), '予約表シートそのものは残る');
m.eq(String(g.__ss.getSheetByName(clsShName).getRange(1, 2).getValue()), '生徒氏名',
  '予約表の見出しも残る');

/* ================================================================
   配布用にまっさらにする（他校へ渡す）
   ================================================================ */
/* 年度末の全初期化では、URL・合言葉・自動処理が残る。
   保護者用URLを消さずに配ると、コピーした学校の案内プリントのQRが
   この学校の予約ページを指し、他校の保護者がこちらへ予約してしまう */

g.setConfigValue_('管理パスコード', 'himitsu');
g.setConfigValue_('保護者用URL', 'https://script.google.com/macros/s/AAA/exec');
g.syncAdminUrl_();
g.dropRefCaches_();
m.ok(String(cfg(g.ADMIN_URL_KEY)).indexOf('page=admin') >= 0, '前提: 管理画面URLが入っている');

g.enableAutoRefresh(5);
g.enableNgAutoApply();
g.enableReminder();
g.enableAutoBackup();
m.eq(triggerCount('refreshViewsIfStale'), 1, '前提: 表示の自動更新がオン');
m.eq(triggerCount('onNgEdit'), 1, '前提: だめなコマの自動反映がオン');
m.eq(triggerCount('dailyReminderCheck'), 1, '前提: 自動リマインドがオン');
m.eq(triggerCount('dailyBackupIfOpen'), 1, '前提: 自動バックアップがオン');

const stopped = g.clearSchoolIdentity_();

m.eq(cfg('保護者用URL'), '',
  '**保護者用URLを空にする**（残すと他校の保護者がこちらへ予約してしまう）');
m.eq(cfg(g.ADMIN_URL_KEY), '', '管理画面URLも空にする');
m.eq(cfg('管理パスコード'), '', '管理パスコードも空にする');

m.eq(stopped, 4, '止めた自動処理の数を返す');
m.eq(triggerCount('refreshViewsIfStale'), 0, '表示の自動更新が止まる');
m.eq(triggerCount('onNgEdit'), 0, 'だめなコマの自動反映が止まる');
m.eq(triggerCount('dailyReminderCheck'), 0, '自動リマインドが止まる');
m.eq(triggerCount('dailyBackupIfOpen'), 0, '自動バックアップが止まる');

const flags = g.readAutoFlags_();
m.eq(flags[g.VIEW_AUTO_KEY], false, '設定シートの表示も、止めた状態に揃う');
m.eq(flags[g.NG_AUTO_KEY], false, 'だめなコマの自動反映も揃う');
m.eq(flags[g.REMINDER_AUTO_KEY], false, '自動リマインドも揃う');
m.eq(flags[g.BACKUP_AUTO_KEY], false, '自動バックアップも揃う');

// ひな形として使えるだけのものは残っている
m.eq(g.getClasses().map((c) => c.name), [C1, C2], 'クラス名は残る');
m.eq(g.getConfig().slotsPerDay, 3, '面談の時間設定は残る');

/* ---------------- URLが空なら、管理画面URLを復活させない ---------------- */
/* ひな形を配ったあとに誰かが表を更新しても、こちらのURLが戻ってはいけない */

g.syncAdminUrl_();
m.eq(cfg(g.ADMIN_URL_KEY), '',
  '保護者用URLが空のうちは、管理画面URLを書き戻さない');

g.dropRefCaches_();
m.throwsWith(() => g.exportHandoutPdf(), '「保護者用URL」が空欄です',
  'URLが空のまま案内プリントを作らせない（誤ったQRを紙に残さない）');

m.report('test_reset');
