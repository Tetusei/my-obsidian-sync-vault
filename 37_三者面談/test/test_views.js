/**
 * 予約の確定と、表示の更新を切り離した仕組み。
 *
 * 予約が1件入るたびに全クラスの表を作り直していたころ、17クラスでは
 * 1件あたり1分以上かかり、排他ロックのぶんそのまま保護者の待ち行列になっていた。
 *
 * いまは予約時に「表が古くなった」印だけを付け、作り直しは数分おきにまとめて行う。
 * **表示が最大5分遅れても、予約できる・できないの判定は常に最新。**
 * この約束が崩れると二重予約が起きるので、そこを重点的に確かめる。
 */

'use strict';

const m = require('./gasmock');

const g = m.load();
m.seedSchool(g, { perClass: 3 });

const C1 = '1年1組';
const DAY1 = '2026-10-29';

const A = { cls: C1, no: 1, name: '生徒1_1' };
const B = { cls: C1, no: 2, name: '生徒1_2' };
const s1 = m.slotIdOf(g, DAY1, C1, 1);
const s2 = m.slotIdOf(g, DAY1, C1, 2);

const clsSh = () => g.__ss.getSheetByName(g.CLASS_SHEET_PREFIX + C1);

/** 「設定」シートを人が書き換えたことにして、取付トリガーを呼ぶ */
function editConfig(key, value) {
  const sh = g.__ss.getSheetByName(g.SH.CONFIG);
  const last = sh.getLastRow();
  const keys = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < keys.length; i++) {
    if (String(keys[i][0]).trim() !== key) continue;
    const range = sh.getRange(i + 2, 2);
    range.setValue(value);
    g.dropRefCaches_();
    g.onConfigEdit({ range: range, value: value });
    return true;
  }
  return false;
}

function triggerCount(fn) {
  return g.__state.triggers.filter((t) => t.getHandlerFunction() === fn).length;
}

/* ---------------- 印を付けるだけで、表は触らない ---------------- */

g.refreshViews(true);
m.eq(g.pendingViewUpdates_(), 0, '作り直した直後は、未反映が0');

m.unwrap(g.apiBook(Object.assign({}, A, { slotId: s1 })), 'Aが予約');
m.eq(g.pendingViewUpdates_(), 1, '予約すると「表が古くなった」印が1つ付く');

m.eq(String(clsSh().getRange(2, 3).getValue()), '未予約',
  '**表はまだ古いまま**（作り直していないので）');
m.eq(m.statusOf(g, s1), '予約済', 'ただし枠マスタは、その場で予約済になっている');

/* ---------------- 表示が古くても、判定は最新 ---------------- */
/* ここが崩れると二重予約になる。性能のために表示を遅らせている以上、
   判定だけは絶対に枠マスタを見ていなければならない */

m.failsWith(g.apiBook(Object.assign({}, B, { slotId: s1 })),
  'ちょうど埋まりました',
  '**表が古いままでも、埋まった枠は別の生徒に取らせない**');

const view = m.unwrap(g.apiSlots(B), '保護者の一覧');
const cell = view.days[0].slots.find((s) => s.slotId === s1);
m.ok(!cell.available, '保護者の画面でも、その枠は選べない');

m.unwrap(g.apiLookup(Object.assign({}, A, { code: 'x' })).ok ? {} : { ok: true }, 'ダミー');
m.eq(g.pendingViewUpdates_(), 1, '失敗した予約では、印は増えない');

/* ---------------- 変更・取消でも印が付く ---------------- */

const booked = m.unwrap(g.apiLookup(Object.assign({}, A,
  { code: m.slotValue(g, s1, g.COL.CODE) })), '予約コードを拾う');
const code = m.slotValue(g, s1, g.COL.CODE);

m.unwrap(g.apiChange(Object.assign({}, A, { code: code, slotId: s2 })), '変更');
m.eq(g.pendingViewUpdates_(), 2, '変更でも印が付く');

m.unwrap(g.apiCancel(Object.assign({}, A, { code: code })), '取消');
m.eq(g.pendingViewUpdates_(), 3, '取消でも印が付く');

/* ---------------- まとめて反映する ---------------- */

const res = g.refreshViews(false);
m.ok(res.refreshed, '未反映があれば作り直す');
m.eq(res.pending, 3, '何件たまっていたかを返す');
m.eq(g.pendingViewUpdates_(), 0, '作り直したら印は消える');

const skipped = g.refreshViews(false);
m.eq(skipped.refreshed, false, '未反映が無ければ、作り直さない（5分おきに走るので無駄を避ける）');
m.eq(skipped.pending, 0, '0件と返す');

m.ok(g.refreshViews(true).refreshed, '未反映が無くても、force なら作り直す');

const at = g.viewsUpdatedAt_();
m.ok(at instanceof Date, '最後に作り直した時刻を覚える');
m.ok(Date.now() - at.getTime() < 60000, 'いま作り直した時刻になっている');

/* ---------------- 作り直すと、表が追いつく ---------------- */

const b2 = m.unwrap(g.apiBook(Object.assign({}, B, { slotId: s2, guardian: '保護者B' })), 'Bが予約');
m.eq(String(clsSh().getRange(3, 3).getValue()), '未予約', '作り直す前は古いまま');

g.refreshViews(false);
m.eq(String(clsSh().getRange(3, 3).getValue()), '予約済', '作り直すと表が追いつく');
m.ok(String(clsSh().getRange(3, 4).getValue()).indexOf('14:05') >= 0, '日時も入る');
m.eq(String(clsSh().getRange(3, 5).getValue()), '保護者B', '保護者氏名も入る');

const ovSh = g.__ss.getSheetByName(g.SH.OVERVIEW);
m.ok(String(ovSh.getRange(1, 1).getValue()).indexOf('1/' + g.getRoster().length + '名') >= 0,
  '全体ビューの進捗（○/○名）も追いつく');

/* ---------------- 作り直しの途中に入った予約を取りこぼさない ---------------- */
/* 印は「作り直す前」に消す。あとで消すと、作り直しているあいだに
   入った予約の印まで一緒に消えてしまう */

m.unwrap(g.apiCancel(Object.assign({}, B, { code: b2.code })), '後片付け');
g.refreshViews(true);
m.eq(g.pendingViewUpdates_(), 0, '前提: 未反映は0');

// 作り直しの最中に予約が入った状況を、表の作り直しに割り込んで作る
const origRebuild = g.rebuildClassSheets;
let injected = null;
g.rebuildClassSheets = function () {
  if (!injected) {
    injected = m.unwrap(g.apiBook(Object.assign({}, A, { slotId: s1 })), '作り直しの最中の予約');
  }
  return origRebuild.apply(null, arguments);
};
g.refreshViews(true);
g.rebuildClassSheets = origRebuild;

m.eq(g.pendingViewUpdates_(), 1,
  '**作り直している最中に入った予約の印は、消されずに残る**（次の更新で反映される）');
g.refreshViews(false);
m.unwrap(g.apiCancel(Object.assign({}, A, { code: injected.code })), '後片付け');
g.refreshViews(true);

/* ---------------- 自動更新のタイマー ---------------- */

m.eq(triggerCount('refreshViewsIfStale'), 0, '最初はタイマーが無い');
m.eq(g.autoRefreshEnabled_(), false, 'オフと判定する');

const minutes = g.enableAutoRefresh(5);
m.eq(minutes, 5, '5分おきで設定する');
m.eq(triggerCount('refreshViewsIfStale'), 1, 'タイマーが1つ付く');
m.ok(g.autoRefreshEnabled_(), 'オンと判定する');

g.enableAutoRefresh(10);
m.eq(triggerCount('refreshViewsIfStale'), 1,
  '設定し直しても、タイマーが二重にならない（先に消してから付ける）');

m.eq(g.enableAutoRefresh(7), 5, '選べない間隔を渡したら5分にする');
m.eq(g.enableAutoRefresh(30), 30, '30分は選べる');

/* ---------------- 設定シートに実態を書き写す ---------------- */
/* メニューを組み立てる onOpen は認可が無く、トリガーを見に行けない。
   そのため実態をシートへ書き写し、メニューの表示はそちらを読む */

const flags = g.readAutoFlags_();
m.eq(flags[g.VIEW_AUTO_KEY], true, 'オンの状態が設定シートに書き写される');

g.disableAutoRefresh();
m.eq(triggerCount('refreshViewsIfStale'), 0, 'タイマーが消える');
m.eq(g.readAutoFlags_()[g.VIEW_AUTO_KEY], false, 'オフも書き写される');

/* ---------------- 設定シートから切り替えられる ---------------- */

editConfig(g.VIEW_AUTO_KEY, true);
m.eq(triggerCount('refreshViewsIfStale'), 1,
  '設定シートを TRUE にすると、タイマーが実際に付く');

editConfig(g.VIEW_AUTO_KEY, false);
m.eq(triggerCount('refreshViewsIfStale'), 0,
  '設定シートを FALSE にすると、タイマーが外れる');

// 同じ値を書き直しても、付け外しを繰り返さない（自分自身を呼び続けないため）
editConfig(g.VIEW_AUTO_KEY, false);
m.eq(triggerCount('refreshViewsIfStale'), 0, '変わらない値では何もしない');

/* ---------------- 枠マスタを直接直したときも印を付ける ---------------- */

const slotSh = g.__ss.getSheetByName(g.SH.SLOTS);
g.refreshViews(true);
m.eq(g.pendingViewUpdates_(), 0, '前提: 未反映は0');

const range = slotSh.getRange(2, g.COL.NOTE);
range.setValue('手で直した');
g.onConfigEdit({ range: range, value: '手で直した' });
m.eq(g.pendingViewUpdates_(), 1,
  '枠マスタは元データなので、手で直しても表に反映されるよう印を付ける');

/* ---------------- 受付中にオフだと、点検で知らせる ---------------- */

g.disableAutoRefresh();
g.dropRefCaches_();
const check = g.checkData();
const warn = check.warns.find((w) => w.title.indexOf('表示の自動更新がオフです') >= 0);
m.ok(warn, '受付中に自動更新がオフなら、データ点検で知らせる');
m.ok(warn.detail.indexOf('1 件') >= 0, 'たまっている未反映の件数も出す');
m.ok(warn.fix.indexOf('オフのままでも予約は正しく処理されます') >= 0,
  '慌てさせないよう、予約自体は無事だと伝える');

g.enableAutoRefresh(5);
g.dropRefCaches_();
m.ok(!g.checkData().warns.some((w) => w.title.indexOf('表示の自動更新がオフです') >= 0),
  'オンにすれば、点検から消える');

/* ---------------- タイマーからの呼び出しは、失敗しても止まらない ---------------- */

const origOverview = g.rebuildOverview;
g.rebuildOverview = function () { throw new Error('わざと失敗'); };
g.refreshViewsIfStale();          // 例外を投げずに戻ってくること
g.rebuildOverview = origOverview;
m.ok(true, 'タイマーからの作り直しが失敗しても、例外を投げて止まらない');

m.report('test_views');
