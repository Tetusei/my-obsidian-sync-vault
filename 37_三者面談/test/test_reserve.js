/**
 * 予備コマ。
 *
 * 保護者には見せず、担任が「もしものとき」に手で埋める枠。
 * v4.6.13〜16 で入り口も取り込みも作り替えたばかりで、
 * そのときバグが2つ出ている（入れた生徒が次の更新で消える／手入力が素通し）。
 *
 * 入り口は2つだけ。
 *   ・担任用の管理画面の黄色いセル
 *   ・「予約表_〇組」で行を選んで、メニューの「この枠に生徒を入れる」
 * **表に直接書いても登録されない。**
 */

'use strict';

const m = require('./gasmock');

const g = m.load();
m.seedSchool(g, { perClass: 3 });

const CLS = '1年1組';
const DAY1 = '2026-10-29';
const PASS = 'test1234';

g.setConfigValue_('管理パスコード', PASS);
g.setConfigValue_(g.RESERVE_COUNT_KEY, 1);
g.dropRefCaches_();
g.generateSlots();

const reserveId = m.slotIdOf(g, DAY1, CLS, g.RESERVE_INDEX_BASE + 1);
const normalId = m.slotIdOf(g, DAY1, CLS, 1);
const reserve2 = m.slotIdOf(g, '2026-10-30', CLS, g.RESERVE_INDEX_BASE + 1);

const S1 = { cls: CLS, no: 1, name: '生徒1_1' };
const S3 = { cls: CLS, no: 3, name: '生徒1_3' };

/** 「予約表_〇組」右の表で、枠IDの行を探す */
function rightRowOf(cls, slotId) {
  const sh = g.__ss.getSheetByName(g.CLASS_SHEET_PREFIX + cls);
  const idCol = 9 + g.CLASS_HEADER_RIGHT.length;
  const last = sh.getLastRow();
  const ids = sh.getRange(2, idCol, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === slotId) return i + 2;
  }
  return 0;
}

/* ---------------- 作られ方 ---------------- */

m.eq(m.statusOf(g, reserveId), '予備', '予備の枠は「予備」で作られる');
m.eq(m.slotValue(g, reserveId, g.COL.START), '14:55',
  '予備は最終コマ(14:30-14:45)の後ろに、同じ休憩をはさんで作られる');
m.eq(m.slotValue(g, reserveId, g.COL.END), '15:10', '長さは通常のコマと同じ');

/* ---------------- 保護者には見せない ---------------- */

const view = m.unwrap(g.apiSlots(S3), '保護者の一覧');
const shown = [];
view.days.forEach((d) => d.slots.forEach((s) => shown.push(s.slotId)));
m.ok(shown.indexOf(reserveId) < 0, '予備の枠は一覧に出さない（存在に気づかせない）');
m.eq(view.existing, null, 'まだ予約は無い');

/* ---------------- シートの小窓から入れる（v4.6.13） ---------------- */

const info = m.unwrap(g.dialogAssignInfo(reserveId), 'dialogAssignInfo');
m.ok(info.reserve, '予備の枠であることを小窓へ伝える（予約コードを出さないため）');
m.eq(info.cls, CLS, 'クラスを返す');
m.ok(info.when.indexOf('14:55') >= 0, '日時を返す');
m.eq(info.students.length, 3, 'そのクラスの名簿を返す（氏名を打たせない）');
m.eq(info.currentName, '', 'まだ誰も入っていない');

const put = m.unwrap(g.dialogAssign(reserveId, '3'), 'dialogAssign で予備に入れる');
m.ok(put.reserve, '予備として処理したことを返す');
m.eq(put.name, '生徒1_3', '入れた生徒の氏名を返す');
m.eq(put.code, undefined, '予備なので予約コードは発行しない');

m.eq(m.statusOf(g, reserveId), '予備', '入れても状態は「予備」のまま');
m.eq(m.slotValue(g, reserveId, g.COL.STUDENT), '生徒1_3', '枠マスタに生徒が入る');
m.eq(m.slotValue(g, reserveId, g.COL.NUMBER), '3', '出席番号も入る');
m.eq(m.slotValue(g, reserveId, g.COL.CODE), '',
  '予備に予約コードは無い（保護者の画面から辿り着けない枠なので）');

/* ---------------- 表を作り直しても消えない（v4.6.13 の不具合） ---------------- */
/* dialogAssign は中で表を作り直す。かつて取り込み処理が、
   まだ空のままの表を正として枠マスタへ書き戻し、入れた生徒を消していた */

m.eq(m.slotValue(g, reserveId, g.COL.STUDENT), '生徒1_3',
  '**登録直後の作り直しで消えない**');

g.refreshViews(true);
m.eq(m.slotValue(g, reserveId, g.COL.STUDENT), '生徒1_3',
  '**もう一度作り直しても消えない**');

g.refreshViews(true);
g.refreshViews(true);
m.eq(m.slotValue(g, reserveId, g.COL.STUDENT), '生徒1_3',
  '何度作り直しても消えない');

/* ---------------- 予備に入った子は「予約あり」 ---------------- */

m.ok(g.isTakenSlot_(m.slotRow(g, reserveId).v), '予備＋氏名は「面談が入っている」と数える');

const view3 = m.unwrap(g.apiSlots(S3), '予備に入ったあとの保護者の一覧');
m.ok(view3.existing !== null, '保護者には「すでに予約が入っています」と伝わる');
m.eq(view3.existing.slotId, reserveId, '予備の枠が existing に入る');

m.failsWith(g.apiBook(Object.assign({}, S3, { slotId: m.slotIdOf(g, DAY1, CLS, 2) })),
  'すでに予約が入っています', '予備に入っている子は、自分でもう1つ予約できない');

const groups = g.unbookedByClass_();
const grp = groups.find((x) => x.cls === CLS);
m.ok(!grp.students.some((s) => s.no === 3), '未予約の一覧に出さない（担任への催促を防ぐ）');
m.eq(g.countBookings_(), 1, 'リセット画面の件数にも数える');

/* ---------------- 小窓から空にできる（v4.6.15） ---------------- */
/* 表に直接書けなくしたので、外す手段をここに用意していないと詰む */

const info2 = m.unwrap(g.dialogAssignInfo(reserveId), '入っている状態の dialogAssignInfo');
m.eq(info2.currentName, '生徒1_3', 'いま入っている生徒を返す');
m.eq(String(info2.currentNo), '3', '出席番号も返す（プルダウンの初期選択に使う）');

const cleared = m.unwrap(g.dialogAssign(reserveId, ''), 'dialogAssign で空にする');
m.ok(cleared.reserve, '予備として処理する');
m.eq(cleared.name, '', '空にしたことを返す');
m.eq(m.slotValue(g, reserveId, g.COL.STUDENT), '', '枠マスタが空になる');
m.eq(m.statusOf(g, reserveId), '予備', '空にしても枠そのものは残る');

/* ---------------- 通常の枠は代理登録（予約コードが出る） ---------------- */

const ninfo = m.unwrap(g.dialogAssignInfo(normalId), '通常の枠の dialogAssignInfo');
m.ok(!ninfo.reserve, '通常の枠では reserve が立たない');

const assigned = m.unwrap(g.dialogAssign(normalId, '1'), '通常の枠へ代理登録');
m.ok(!assigned.reserve, '通常の枠として処理する');
m.ok(/^\d{4}$/.test(assigned.code), '予約コードを発行する（保護者に伝えるため）');
m.eq(m.statusOf(g, normalId), '予約済', '通常の枠は「予約済」になる');
m.eq(m.slotValue(g, normalId, g.COL.NOTE), '担任が代理で登録', '代理登録と分かるようにする');

// 発行したコードで、保護者があとから変更・取消できること
const look = m.unwrap(g.apiLookup(Object.assign({}, S1, { code: assigned.code })),
  '代理登録のコードで保護者が引ける');
m.eq(look.booking.slotId, normalId, '保護者は自分で手続きできる');

m.unwrap(g.apiCancel(Object.assign({}, S1, { code: assigned.code })), '後片付け');

/* ---------------- 管理画面から ---------------- */

m.failsWith(g.apiAdminSetReserve('ちがう合言葉', reserveId, '3'),
  'パスコードが違います', '合言葉が違えば触れない');

const web = m.unwrap(g.apiAdminSetReserve(PASS, reserveId, '生徒1_3'), '管理画面から氏名で入れる');
m.eq(web.name, '生徒1_3', '氏名でも入れられる');
m.eq(web.no, 3, '出席番号を名簿から補う');

m.unwrap(g.apiAdminSetReserve(PASS, reserveId, ''), '管理画面から空にする');
m.eq(m.slotValue(g, reserveId, g.COL.STUDENT), '', '空になる');

/* ---------------- 弾くべきもの ---------------- */

m.throwsWith(() => g.setReserveStudent_(normalId, '1', 'テスト'),
  '予備の枠ではありません', '通常の枠を予備として扱わせない');

m.throwsWith(() => g.assignStudentToSlot_(reserveId, '1'),
  '予備の枠です', '予備の枠を代理登録の経路へ通さない');
m.throwsWith(() => g.assignStudentToSlot_(reserveId, '1'),
  '管理画面の黄色いセル', '正しい入り口を案内する');

m.failsWith(g.apiAdminSetStatus(PASS, reserveId, 'ブロック'),
  '予備の枠はここでは変更できません', '予備の枠は空き／ブロックに切り替えられない');

m.throwsWith(() => g.setReserveStudent_(reserveId, 'いない子', 'テスト'),
  '名簿に「いない子」が見つかりません', '名簿に無い氏名を弾く');
m.throwsWith(() => g.setReserveStudent_(reserveId, '99', 'テスト'),
  '名簿に「99」が見つかりません', '名簿に無い出席番号を弾く');

/* ---------------- 二重に入れない ---------------- */

const booked = m.unwrap(g.apiBook(Object.assign({}, S1, { slotId: normalId })), 'S1が自分で予約');
m.throwsWith(() => g.setReserveStudent_(reserveId, '1', 'テスト'),
  '面談が入っています', 'すでに面談がある生徒を、予備にも入れない');
m.eq(m.slotValue(g, reserveId, g.COL.STUDENT), '', '弾いたので予備は空のまま');

m.unwrap(g.apiCancel(Object.assign({}, S1, { code: booked.code })), '取消');

g.setReserveStudent_(reserveId, '1', 'テスト');
m.throwsWith(() => g.setReserveStudent_(reserve2, '1', 'テスト'),
  '面談が入っています', '同じ生徒を2つの予備に入れない');

// 入っている本人を入れ直すのは通す（同じ枠なので二重にならない）
g.setReserveStudent_(reserveId, '1', 'テスト');
m.eq(m.slotValue(g, reserveId, g.COL.STUDENT), '生徒1_1', '同じ枠へ入れ直すのは通る');

/* ---------------- 表への手入力は取り込まない（v4.6.15） ---------------- */

g.refreshViews(true);
const row = rightRowOf(CLS, reserve2);
m.ok(row > 0, '予約表の右の表に、予備の行がある');

const sh = g.__ss.getSheetByName(g.CLASS_SHEET_PREFIX + CLS);
sh.getRange(row, 12, 1, 2).setValues([[2, '生徒1_2']]);   // 出席番号・生徒氏名に直接書く

g.refreshViews(true);
m.eq(m.slotValue(g, reserve2, g.COL.STUDENT), '',
  '**表に直接書いても、枠マスタには取り込まれない**');
m.eq(String(sh.getRange(rightRowOf(CLS, reserve2), 13).getValue()), '',
  '手で書いた内容は、次の作り直しで消える');

/* ---------------- 表示（v4.6.16 の色） ---------------- */

g.setReserveStudent_(reserveId, '3', 'テスト');       // 3番は予備
m.unwrap(g.apiBook(Object.assign({}, S1, { slotId: m.slotIdOf(g, DAY1, CLS, 2) })), '1番は通常の枠');
g.refreshViews(true);                                  // 2番は未予約のまま

const left = sh.getRange(2, 1, 3, 3).getValues();
const bg = sh.getRange(2, 1, 3, 3).getBackgrounds();

m.eq(left.map((r) => String(r[0])), ['1', '2', '3'], '出席番号順に並ぶ');
m.eq(String(left[0][2]), '予約済', '1番は予約済');
m.eq(String(left[1][2]), '未予約', '2番は未予約');
m.eq(String(left[2][2]), '予約済（予備）', '3番は予約済（予備）');

m.eq(bg[0][2], '#e6f4ea', '予約済は緑');
m.eq(bg[1][2], '#fef7e0', '未予約は琥珀');
m.eq(bg[2][2], '#ceead6', '予約済（予備）は一段濃い緑');

m.ok(bg[2][2] !== bg[1][2],
  '**予約済（予備）と未予約は違う色**（同じだと埋まったのが分からない）');
m.ok(bg[2][2] !== bg[0][2], '通常の予約とも見分けが付く');

/* ---------------- 一括取消では、枠そのものは残す ---------------- */

g.clearAllBookings_();
m.eq(m.statusOf(g, reserveId), '予備', '一括取消しても予備の枠は残る');
m.eq(m.slotValue(g, reserveId, g.COL.STUDENT), '', '中の生徒だけ消える');
m.eq(m.statusOf(g, m.slotIdOf(g, DAY1, CLS, 2)), '空き', '通常の枠は空きに戻る');

/* ---------------- ログ ---------------- */

const log = g.__ss.getSheetByName(g.SH.LOG);
const rows = log.getRange(2, 1, log.getLastRow() - 1, 7).getValues();
const actions = rows.map((r) => String(r[1]));
m.ok(actions.indexOf('予備コマに記入') >= 0, '予備への記入がログに残る');
m.ok(actions.indexOf('予備コマを空に') >= 0, '予備を空にしたこともログに残る');

const fromSheet = rows.filter((r) => String(r[6]) === '担任シート');
const fromWeb = rows.filter((r) => String(r[6]) === '担任Web');
m.ok(fromSheet.length > 0, 'シートの小窓からの操作と分かる');
m.ok(fromWeb.length > 0, '管理画面からの操作と分かる');

m.report('test_reserve');
