/**
 * 作り直される場所への手入力を止める仕掛け（v4.6.14〜15）。
 *
 * 手引きを読まず、感覚だけでシートを操作する担任は必ずいる。
 * 予約表や全体ビューへ手で書いても次の更新で消えるが、
 * **その時間は保護者の画面では空きのままなので、担任は「入れたつもり」、
 * 保護者は別の家庭が予約できる状態**になる。当日まで誰も気づかない。
 *
 * 二段構えで塞ぐ。
 *   1. 書き込む瞬間に、Googleの保護（警告のみ）で止める
 *   2. それでも書かれたら、予約ログに残してデータ点検から拾う
 *
 * 手で入れてよいのは A・B列の名簿だけ、と一言で言い切れる状態を保つ。
 */

'use strict';

const m = require('./gasmock');

const g = m.load();
m.seedSchool(g, { perClass: 3 });

const C1 = '1年1組';
const DAY1 = '2026-10-29';

g.setConfigValue_(g.RESERVE_COUNT_KEY, 1);
g.dropRefCaches_();
g.generateSlots();
g.refreshViews(true);

const clsSh = () => g.__ss.getSheetByName(g.CLASS_SHEET_PREFIX + C1);
const ovSh = () => g.__ss.getSheetByName(g.SH.OVERVIEW);

/** 目印の付いた保護だけを集める */
function guards(sh) {
  return sh.getProtections('RANGE').concat(sh.getProtections('SHEET'))
    .filter((p) => String(p.getDescription()).indexOf(g.EDIT_GUARD_TAG) === 0);
}

/**
 * 人がシートを編集したことにして onEdit を呼ぶ。
 * value に配列を渡すと「複数のセルへ貼り付けた」ことになる。
 * 実機と同じく、そのときイベントの value は空になる（1セルのときだけ入る）。
 */
function edit(sheetName, row, col, value, numCols) {
  const sh = g.__ss.getSheetByName(sheetName);
  const many = Array.isArray(value);
  const range = sh.getRange(row, col, 1, many ? value.length : (numCols || 1));
  if (many) range.setValues([value]); else range.setValue(value);
  g.__ss.toasts.length = 0;
  g.onEdit({ range: range, value: many ? undefined : value });
  return g.__ss.toasts;
}

/** 予約ログから「未登録の手入力」だけを拾う */
function strayLogs() {
  const log = g.__ss.getSheetByName(g.SH.LOG);
  const last = log.getLastRow();
  if (last < 2) return [];
  return log.getRange(2, 1, last - 1, 7).getValues()
    .filter((r) => String(r[1]) === g.STRAY_EDIT_ACTION);
}

/* ================================================================
   1. 書き込む瞬間に止める（保護：警告のみ）
   ================================================================ */

const cls = guards(clsSh());
m.eq(cls.length, 2, '予約表には2つの保護を掛ける（左の表と右の表）');
m.ok(cls.every((p) => p.isWarningOnly()),
  '**警告のみ**にする。［OK］で書けるので作業は止まらない');

const byKey = {};
cls.forEach((p) => { byKey[p.getRange().getA1Notation()] = p.getDescription(); });

m.ok(byKey['C2:F1000'], '左の表（C〜F列＝予約状況・予約日時・保護者氏名・連絡事項）を守る');
m.ok(byKey['I2:P1000'], '右の表（I〜P列）をまるごと守る。予備の行も含む');

m.ok(String(byKey['C2:F1000']).indexOf('A・B列') >= 0,
  '説明に、どこへ書けばよいかを入れる');
m.ok(String(byKey['I2:P1000']).indexOf('予備') >= 0,
  '右の表の説明で、予備の行も同じだと明示する');
m.ok(String(byKey['I2:P1000']).indexOf('この枠に生徒を入れる') >= 0,
  '説明に、正しい入り口を書く');

// A・B列（名簿）は手で入れる場所なので、守らない
const a1Guarded = cls.some((p) => {
  const r = p.getRange();
  return r.getColumn() <= 2;
});
m.ok(!a1Guarded, '**A・B列（名簿）には保護を掛けない**（正しい操作に警告を出さない）');

const ov = guards(ovSh());
m.eq(ov.length, 1, '全体ビューは1つ');
m.eq(ov[0].getRange(), null, '全体ビューは clear() で作り直すので、シートごと守る');
m.ok(ov[0].isWarningOnly(), '全体ビューも警告のみ');

/* ---------------- 何度作り直しても増えない ---------------- */

g.refreshViews(true);
g.refreshViews(true);
m.eq(guards(clsSh()).length, 2, '作り直しても保護は増えない（5分おきに走るので）');
m.eq(guards(ovSh()).length, 1, '全体ビューも増えない');

/* ---------------- 二重に付いたら片付ける ---------------- */

// 更新が同時に2つ走ると、同じ保護が二重にできることがある
clsSh().getRange(2, 3, 999, 4).protect()
  .setDescription(g.EDIT_GUARD_TAG + '／生徒別予約状況：二重にできたもの')
  .setWarningOnly(true);
m.eq(guards(clsSh()).length, 3, '二重になった状態を作る');

g.refreshViews(true);
m.eq(guards(clsSh()).length, 2, '次の作り直しで、二重ぶんを片付ける');

/* ---------------- 使わなくなった保護は外す ---------------- */

clsSh().getRange(2, 1, 999, 2).protect()
  .setDescription(g.EDIT_GUARD_TAG + '／古い版のキー：もう使っていない範囲')
  .setWarningOnly(true);
g.refreshViews(true);
m.ok(!guards(clsSh()).some((p) => String(p.getDescription()).indexOf('古い版のキー') >= 0),
  '守る場所を変えたら、前の版の保護は外す（いらない所に警告が出続けないように）');
m.eq(guards(clsSh()).length, 2, '外したあとも、いまの2つは残る');

/* ---------------- 学校が自分で付けた保護には触らない ---------------- */

clsSh().getRange(1, 1, 1, 6).protect().setDescription('教務主任が付けた保護');
g.refreshViews(true);
const mine = clsSh().getProtections('RANGE')
  .filter((p) => p.getDescription() === '教務主任が付けた保護');
m.eq(mine.length, 1, '目印の付いていない保護は、システムが勝手に外さない');

/* ================================================================
   2. すり抜けたら、必ず記録に残す
   ================================================================ */

const before = strayLogs().length;

/* ---------------- 予約表の、作り直される場所 ---------------- */

let toasts = edit(g.CLASS_SHEET_PREFIX + C1, 2, 4, '10月29日 13:40');
m.eq(toasts.length, 1, '書いた瞬間にトーストで知らせる');
m.ok(toasts[0].title.indexOf('残りません') >= 0, 'タイトルで「残らない」と伝える');
m.ok(toasts[0].msg.indexOf('A・B列の名簿だけ') >= 0, 'どこへ書けばよいかを伝える');
m.ok(toasts[0].msg.indexOf('この枠に生徒を入れる') >= 0, '正しい入り口を伝える');

let logs = strayLogs();
m.eq(logs.length, before + 1, 'トーストは15秒で消えるので、予約ログにも残す');
m.eq(String(logs[logs.length - 1][3]), C1, 'どのクラスかを残す');
m.ok(String(logs[logs.length - 1][6]).indexOf('D2') >= 0, 'どのセルかを残す');
m.ok(String(logs[logs.length - 1][6]).indexOf('10月29日 13:40') >= 0, '書かれた内容も残す');

/* ---------------- 予備の行も同じ（v4.6.15） ---------------- */
/* v4.6.14 までは予備の行だけ手入力を許していたが、
   「ここはよくて、ここはだめ」は覚えられないので全部だめに統一した */

const reserveId = m.slotIdOf(g, DAY1, C1, g.RESERVE_INDEX_BASE + 1);
const idCol = 9 + g.CLASS_HEADER_RIGHT.length;
const ids = clsSh().getRange(2, idCol, clsSh().getLastRow() - 1, 1).getValues();
let reserveRow = 0;
for (let i = 0; i < ids.length; i++) {
  if (String(ids[i][0]).trim() === reserveId) { reserveRow = i + 2; break; }
}
m.ok(reserveRow > 0, '予約表に予備の行がある');
m.eq(String(clsSh().getRange(reserveRow, 11).getValue()), '予備', 'その行は予備');

toasts = edit(g.CLASS_SHEET_PREFIX + C1, reserveRow, 13, '生徒1_2');
m.eq(toasts.length, 1, '**予備の行に書いても警告する**（例外を作らない）');
m.eq(strayLogs().length, before + 2, '予備の行の手入力もログに残る');

/* ---------------- 名簿（A・B列）は何も言わない ---------------- */

toasts = edit(g.CLASS_SHEET_PREFIX + C1, 5, 1, 4);
m.eq(toasts.length, 0, 'A列（出席番号）は正しい入力先なので、警告しない');
toasts = edit(g.CLASS_SHEET_PREFIX + C1, 5, 2, '転入 太郎');
m.eq(toasts.length, 0, 'B列（生徒氏名）も警告しない');
toasts = edit(g.CLASS_SHEET_PREFIX + C1, 5, 1, [4, '転入 太郎']);
m.eq(toasts.length, 0, 'A・B列をまとめて貼っても警告しない');
m.eq(strayLogs().length, before + 2, '名簿の入力はログにも残さない');

toasts = edit(g.CLASS_SHEET_PREFIX + C1, 1, 3, '見出し');
m.eq(toasts.length, 0, '見出しの行（1行目）は対象外');

/* ---------------- 全体ビュー（v4.6.14 で追加） ---------------- */

toasts = edit(g.SH.OVERVIEW, 5, 4, 'メモ');
m.eq(toasts.length, 1, '全体ビューへ書いても知らせる（v4.6.14 まで警告すら出なかった）');
m.ok(toasts[0].msg.indexOf(g.SH.NG) >= 0,
  '面談を入れない指定は、だめなコマシートか管理画面だと案内する');

logs = strayLogs();
m.eq(logs.length, before + 3, '全体ビューの手入力もログに残る');
m.eq(String(logs[logs.length - 1][3]), '', '全体ビューはクラスが無いので空欄');

/* ---------------- 手で入れる前提のシートは対象外 ---------------- */

toasts = edit(g.SH.CLASSES, 2, 2, '新担任');
m.eq(toasts.length, 0, 'クラスシートは手で入れる場所なので警告しない');
toasts = edit(g.SH.DAYS, 2, 2, '備考');
m.eq(toasts.length, 0, '面談日シートも警告しない');
toasts = edit(g.SH.NG, 4, 5, true);
m.eq(toasts.length, 0, 'だめなコマシートは、チェックを入れる場所なので警告しない');

/* ================================================================
   3. データ点検から拾う
   ================================================================ */

const recent = g.recentStrayEdits_(14);
m.ok(recent.length >= 2, 'クラスごとにまとめて拾う');
const c1Rec = recent.find((x) => x.cls === C1);
m.eq(c1Rec.count, 2, C1 + ' の手入力を2件数える');
const ovRec = recent.find((x) => x.cls === g.SH.OVERVIEW);
m.eq(ovRec.count, 1, '全体ビューぶんは「全体ビュー」としてまとめる');

const check = g.checkData();
const found = check.warns.find((w) => w.title.indexOf('手入力がありました') >= 0);
m.ok(found, 'データ点検で「確認しておきたいもの」に出る');
m.ok(found.title.indexOf('3件') >= 0, '合計の件数を出す');
m.ok(found.detail.indexOf(C1) >= 0, 'どのクラスかを出す');
m.ok(found.fix.indexOf('入れたつもり') >= 0, '何が起きているかを説明する');
m.ok(found.fix.indexOf('別の家庭が予約できてしまいます') >= 0,
  '放っておくとどうなるかを伝える（ここが伝わらないと動いてもらえない）');

const pre = g.publishPreflight_();
m.ok(pre.warnings.some((w) => w.indexOf('手入力がありました') >= 0),
  '受付を開始するときの点検にも出る');

/* ---------------- 古い記録は出さない ---------------- */

const log = g.__ss.getSheetByName(g.SH.LOG);
const old = new Date();
old.setDate(old.getDate() - 20);
log.getRange(log.getLastRow(), 1).setValue(old);
m.eq(g.recentStrayEdits_(14).reduce((n, x) => n + x.count, 0), 2,
  '14日より前の記録は数えない（いつまでも警告が残らないように）');

/* ---------------- 手入力は取り込まれない ---------------- */

m.eq(m.slotValue(g, reserveId, g.COL.STUDENT), '',
  '予備の行に書いた内容は、枠マスタへ取り込まれていない');

g.refreshViews(true);
m.eq(String(clsSh().getRange(reserveRow, 13).getValue()), '',
  '書いた内容は、次の作り直しで消える');
m.eq(String(clsSh().getRange(2, 4).getValue()), '—',
  '予約状況の欄も、作り直しで元に戻る');

// A・B列に書いた名簿は残る
m.eq(String(clsSh().getRange(5, 2).getValue()), '転入 太郎',
  '**名簿（A・B列）に書いたものは残る**（ここにしか控えが無いため）');
m.ok(g.getRoster().some((r) => r.name === '転入 太郎'), '名簿として読み込まれる');

m.report('test_guard');
