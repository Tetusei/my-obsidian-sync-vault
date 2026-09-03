/**
 * 予約と「だめなコマ」がぶつかったときの警告。
 *
 * 担任がチェックを入れた時点で、すでに保護者の予約が入っていることがある。
 * その予約は消さない（消したら保護者は来てしまう）。かわりに、
 * **気づけるように、あらゆる画面へ警告を出す。**
 *
 * 警告が出ないと、担任が居ない時間に保護者が来る。
 * 解消したのに警告が残ると、今度は警告そのものが信用されなくなる。
 */

'use strict';

const m = require('./gasmock');

const g = m.load();
m.seedSchool(g, { perClass: 3 });

const C1 = '1年1組';
const DAY1 = '2026-10-29';
const PASS = 'test1234';

g.setConfigValue_('管理パスコード', PASS);
g.dropRefCaches_();

const STU = { cls: C1, no: 1, name: '生徒1_1' };
const slotId = m.slotIdOf(g, DAY1, C1, 2);   // 1日目の2コマ目

const ngSh = () => g.__ss.getSheetByName(g.SH.NG);
const clsSh = () => g.__ss.getSheetByName(g.CLASS_SHEET_PREFIX + C1);
const ovSh = () => g.__ss.getSheetByName(g.SH.OVERVIEW);

/** だめなコマシートの (クラス, 日, コマ) の位置 */
function ngPos(cls, ymd, idx) {
  const sh = ngSh();
  const lastCol = sh.getLastColumn();
  const header = sh.getRange(g.NG_HEADER_ROW, 1, 1, lastCol).getValues()[0];
  let col = 0;
  for (let c = g.NG_COL.FIRST_CLASS - 1; c < lastCol; c++) {
    if (g.ngClassNameOf_(header[c]) === cls) { col = c + 1; break; }
  }
  const last = sh.getLastRow();
  const rows = sh.getRange(g.NG_FIRST_ROW, 1, last - g.NG_FIRST_ROW + 1, 4).getValues();
  let row = 0;
  for (let r = 0; r < rows.length; r++) {
    if (rows[r][0] && g.ymd_(rows[r][0]) === ymd && Number(rows[r][2]) === idx) {
      row = r + g.NG_FIRST_ROW; break;
    }
  }
  return { row, col };
}

/** 全体ビューで、その枠にあたるセル */
function overviewCell(cls, ymd, idx) {
  const sh = ovSh();
  const lastCol = sh.getLastColumn();
  const header = sh.getRange(2, 1, 1, lastCol).getValues()[0];
  // v4.6.17 以降、見出しはクラス予約表へのリンク（=HYPERLINK の数式）になっている。
  // 数式の中にクラス名が入っているので、部分一致で探す
  let col = 0;
  for (let c = 2; c < lastCol; c++) {
    if (String(header[c]).indexOf(cls) >= 0) { col = c + 1; break; }
  }
  const times = g.dayAllTimes_(g.getConfig()).length;
  const days = g.getDays().map((d) => g.ymd_(d));
  const row = 3 + days.indexOf(ymd) * times + (idx - 1);
  return { row, col };
}

/** 予約表の右の表で、枠IDの行 */
function rightRow(slot) {
  const sh = clsSh();
  const idCol = 9 + g.CLASS_HEADER_RIGHT.length;
  const ids = sh.getRange(2, idCol, sh.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]).trim() === slot) return i + 2;
  return 0;
}

/* ---------------- ぶつかりを作る ---------------- */

const booked = m.unwrap(g.apiBook(Object.assign({}, STU, { slotId })), '先に保護者が予約');
m.eq(m.statusOf(g, slotId), '予約済', '予約が入っている');

// あとから担任が「この時間は面談を入れられない」と指定する
const pos = ngPos(C1, DAY1, 2);
ngSh().getRange(pos.row, pos.col).setValue(true);

const res = g.applyNgSlots();
m.eq(res.blocked, 0, '予約済なのでブロックにはしない');
m.eq(res.conflicts.length, 1, 'ぶつかりとして1件数える');

/* ---------------- 何がぶつかっているかを、文字で伝える ---------------- */

const c = res.conflicts[0];
m.eq(c.slotId, slotId, 'どの枠か');
m.eq(c.cls, C1, 'どのクラスか');
m.eq(String(c.no), '1', '出席番号');
m.eq(c.name, '生徒1_1', '生徒氏名');
m.eq(c.code, booked.code, '予約コード（保護者へ連絡するときに要る）');

const label = g.ngConflictLabel_(c);
m.ok(label.indexOf('10月29日') >= 0, '一行の表示に日付が入る');
m.ok(label.indexOf('14:05') >= 0, '時間が入る');
m.ok(label.indexOf('生徒1_1') >= 0, '生徒氏名が入る');

const guide = g.ngWarningText_(res.conflicts);
m.ok(guide.indexOf('1 件') >= 0, '件数を伝える');
m.ok(guide.indexOf('生徒1_1') >= 0, '誰の予約かを伝える');
m.ok(guide.indexOf('保護者に連絡して') >= 0, '何をすればよいかを伝える');
m.ok(guide.indexOf(g.MENU.NG_APPLY) >= 0, 'そのあと何を実行するかを伝える');
m.eq(g.ngWarningText_([]), '', 'ぶつかりが無ければ何も出さない');

/* ---------------- だめなコマシートに、赤とメモを出す ---------------- */

m.eq(ngSh().getRange(pos.row, pos.col).getBackground(), g.NG_CONFLICT_BG,
  'ぶつかっているセルを赤くする');
const note = ngSh().getRange(pos.row, pos.col).getNotes()[0][0];
m.ok(note.indexOf('すでに面談の予約が入っています') >= 0, 'セルのメモで事情を説明する');
m.ok(note.indexOf('生徒1_1') >= 0, 'メモに生徒氏名を入れる');
m.ok(note.indexOf(g.MENU.NG_APPLY) >= 0, 'メモに、解消したあとの手順も入れる');

// 関係のないセルは白のまま
const other = ngPos(C1, DAY1, 1);
m.eq(ngSh().getRange(other.row, other.col).getBackground(), g.NG_NORMAL_BG,
  'ぶつかっていないセルは白のまま');
m.eq(ngSh().getRange(other.row, other.col).getNotes()[0][0], '',
  'ぶつかっていないセルにメモは付けない');

const banner = String(ngSh().getRange(g.NG_BANNER_ROW, 1).getValue());
m.ok(banner.indexOf('⚠') >= 0, '1行目のバナーが警告に切り替わる');
m.ok(banner.indexOf('1 件') >= 0, 'バナーに件数を出す');
m.eq(ngSh().getRange(g.NG_BANNER_ROW, 1).getBackground(), g.NG_CONFLICT_BG,
  'バナーも赤くする');

/* ---------------- 全体ビューに出す ---------------- */

g.refreshViews(true);

const ov = overviewCell(C1, DAY1, 2);
m.ok(String(ovSh().getRange(ov.row, ov.col).getValue()).indexOf('⚠') >= 0,
  '全体ビューのセルに ⚠ を付ける');
m.eq(ovSh().getRange(ov.row, ov.col).getBackground(), '#fce8e6',
  'そのセルを赤くする（学年の色は重ねない。見落とすと困るため）');

const ovBanner = String(ovSh().getRange(1, 3).getValue());
m.ok(ovBanner.indexOf('⚠') >= 0, '全体ビューの上のバナーにも出す');
m.ok(ovBanner.indexOf('1 件') >= 0, 'バナーに件数を出す');
m.ok(ovBanner.indexOf(g.SH.NG) >= 0, 'どのシートを見ればよいかを書く');
m.eq(ovSh().getRange(1, 1).getBackground(), '#fce8e6', 'バナー全体を赤くする');

/* ---------------- クラス別予約表に出す ---------------- */

const leftRow = 2;   // 出席番号1番＝最初の行
m.eq(String(clsSh().getRange(leftRow, 3).getValue()), '⚠ 要調整',
  '生徒別の一覧で「⚠ 要調整」にする');
m.eq(clsSh().getRange(leftRow, 3).getBackground(), '#fce8e6', 'その行を赤くする');
m.ok(String(clsSh().getRange(leftRow, 6).getValue()).indexOf('時間の変更をご相談ください') >= 0,
  '連絡事項の欄に、何をすればよいかを書く');

const rr = rightRow(slotId);
m.ok(rr > 0, '時間枠別の表に、その枠の行がある');
m.eq(String(clsSh().getRange(rr, 11).getValue()), '⚠ 予約済（だめなコマ指定）',
  '時間枠別の表でも、ぶつかっていることが分かる');
m.eq(clsSh().getRange(rr, 11).getBackground(), '#fce8e6', 'その行も赤くする');

/* ---------------- データ点検と管理画面に出す ---------------- */

const check = g.checkData();
const found = check.warns.find((w) => w.title.indexOf('だめなコマに予約が残っています') >= 0);
m.ok(found, 'データ点検で「確認しておきたいもの」に出る');
m.ok(found.detail.indexOf('生徒1_1') >= 0, '点検の結果に生徒氏名が出る');
m.ok(found.fix.indexOf('保護者に連絡して') >= 0, '点検の結果に対処法が出る');

const admin = m.unwrap(g.apiAdminInit(PASS), 'apiAdminInit');
m.eq(admin.conflicts.length, 1, '担任用の管理画面にも渡す');
m.ok(admin.conflicts[0].indexOf('生徒1_1') >= 0, '管理画面にも生徒氏名を渡す');

const board = m.unwrap(g.apiAdminBoard(PASS, DAY1), 'apiAdminBoard');
const boardRow = board.rows.find((r) => r.time.indexOf('14:05') >= 0);
m.ok(boardRow.cells[C1].ng, '管理画面の表で、そのセルに ng の印が立つ');
m.eq(boardRow.cells[C1].status, '予約済', '状態は予約済のまま');

/* ---------------- 受付開始前の点検でも止める ---------------- */

const pre = g.publishPreflight_();
m.ok(pre.warnings.some((w) => w.indexOf('だめなコマに予約が残っています') >= 0),
  '受付を開始するときの点検にも出る');

/* ---------------- 解消すると、警告は自動で消える ---------------- */

m.unwrap(g.apiCancel(Object.assign({}, STU, { code: booked.code })), '保護者の予約を取り消す');
const after = g.applyNgSlots();

m.eq(after.conflicts.length, 0, 'ぶつかりが無くなる');
m.eq(m.statusOf(g, slotId), 'ブロック', '指定どおりブロックになる');

m.eq(ngSh().getRange(pos.row, pos.col).getBackground(), g.NG_NORMAL_BG,
  '赤が消える');
m.eq(ngSh().getRange(pos.row, pos.col).getNotes()[0][0], '', 'メモも消える');
m.eq(String(ngSh().getRange(g.NG_BANNER_ROW, 1).getValue()), g.NG_BANNER_TEXT,
  'バナーが元の案内に戻る');
m.eq(ngSh().getRange(g.NG_BANNER_ROW, 1).getBackground(), '#fef7e0', 'バナーの色も戻る');

g.refreshViews(true);
m.eq(String(ovSh().getRange(ov.row, ov.col).getValue()), '×',
  '全体ビューのセルが「×（面談なし）」になる');
m.ok(String(ovSh().getRange(1, 3).getValue()).indexOf('⚠') < 0,
  '全体ビューのバナーからも警告が消える');
m.eq(String(clsSh().getRange(rr, 11).getValue()), 'ブロック',
  '時間枠別の表も普通のブロック表示に戻る');

m.eq(g.checkData().warns.filter((w) => w.title.indexOf('だめなコマに予約が残って') >= 0).length, 0,
  'データ点検からも消える');
m.eq(m.unwrap(g.apiAdminInit(PASS), 'apiAdminInit').conflicts.length, 0,
  '管理画面からも消える');

/* ---------------- 複数ぶつかっても、全部数える ---------------- */

const b1 = m.unwrap(g.apiBook({ cls: C1, no: 1, name: '生徒1_1', slotId: m.slotIdOf(g, DAY1, C1, 1) }), '1人目');
const b2 = m.unwrap(g.apiBook({ cls: C1, no: 2, name: '生徒1_2', slotId: m.slotIdOf(g, DAY1, C1, 3) }), '2人目');

g.setNgFlags_([
  { slotId: m.slotIdOf(g, DAY1, C1, 1), flag: true },
  { slotId: m.slotIdOf(g, DAY1, C1, 3), flag: true }
]);
const multi = g.applyNgSlots();
m.eq(multi.conflicts.length, 2, '2件ともぶつかりとして数える');
m.ok(g.ngWarningText_(multi.conflicts).indexOf('2 件') >= 0, '件数をまとめて伝える');
m.ok(String(ngSh().getRange(g.NG_BANNER_ROW, 1).getValue()).indexOf('2 件') >= 0,
  'バナーにも2件と出る');

// 片方だけ解消しても、残りは警告が続く
m.unwrap(g.apiCancel({ cls: C1, no: 1, name: '生徒1_1', code: b1.code }), '片方を取消');
const half = g.applyNgSlots();
m.eq(half.conflicts.length, 1, '残った1件は警告が続く');
m.eq(half.conflicts[0].name, '生徒1_2', '残っているほうの生徒を指す');

m.unwrap(g.apiCancel({ cls: C1, no: 2, name: '生徒1_2', code: b2.code }), '残りも取消');
m.eq(g.applyNgSlots().conflicts.length, 0, '全部解消すれば警告は消える');

m.report('test_warn');
